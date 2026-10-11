"""Markid Furniture Payroll - V1 (no tax/pension). Flask + SQLite."""
import os, io, re, json, sqlite3, uuid, hmac, time, socket, secrets, tempfile, threading, urllib.request, urllib.error, urllib.parse, datetime as dt
from functools import wraps
from flask import Flask, g, request, session, redirect, send_file, flash, get_flashed_messages
from markupsafe import escape as E, Markup
from werkzeug.security import generate_password_hash as _gph, check_password_hash
def generate_password_hash(p): return _gph(p, method='pbkdf2:sha256')
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'change-me-in-production')
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE='Lax', SESSION_COOKIE_SECURE=os.environ.get('MARKID_HTTPS') == '1')
@app.before_request
def _csrf():
    if 'csrf' not in session: session['csrf'] = os.urandom(16).hex()
    if request.method == 'POST' and not request.path.startswith('/api/') and request.form.get('_csrf') != session['csrf']:
        return 'Session expired or invalid form. Go back, refresh the page and try again.', 400
VERSION = 'v2.0 · 11 Oct 2026'
DB = os.environ.get('MARKID_DB', os.path.join(os.path.dirname(__file__), 'markid.db'))
MODE = os.environ.get('MARKID_MODE', 'standalone')      # standalone | server (online) | client (desktop app)
SYNC_TOKEN = os.environ.get('SYNC_TOKEN', '')
STATE_DB = os.environ.get('MARKID_STATE') or DB + '.sync'
SCHEMA_VERSION = 1
SYNCING = False
UT = {'p': 'payroll_periods', 'w': 'contract_workers', 'd': 'designs', 'e': 'main_employees', 'i': 'contract_work_items', 'o': 'overtime_records', 'n': 'night_work_records'}
ROLES = ['ADMIN', 'ACCOUNTANT', 'CHECKER', 'APPROVER', 'MANAGER', 'VIEWER']
DEFAULTS = {'hours': '192', 'normal_end': '17:30', 'ot_mult': '1.5', 'sun_mult': '2.0', 'sun_hours': '8', 'night_mult': '1.5', 'absence_divisor': '0'}
LABELS = {'hours': 'Monthly working hours', 'normal_end': 'Normal working end (HH:MM)', 'ot_mult': 'Normal overtime multiplier',
          'sun_mult': 'Sunday multiplier', 'sun_hours': 'Sunday standard hours', 'night_mult': 'Night work multiplier (company defines)',
          'absence_divisor': 'Absence: days per month for day-rate deduction (0 = enter the amount manually)'}
NEXT = {'DRAFT': ('CALCULATED', ('ACCOUNTANT',), 'Calculate'), 'CALCULATED': ('CHECKED', ('CHECKER',), 'Mark checked'),
        'CHECKED': ('APPROVED', ('APPROVER',), 'Approve'), 'APPROVED': ('PAID', ('ACCOUNTANT',), 'Mark paid'),
        'PAID': ('LOCKED', ('APPROVER',), 'Lock')}
DIV = {'WEEKLY': 4, 'BIWEEKLY': 2, 'MONTHLY': 1}
RAW = ['transport_records', 'commission_records', 'absence_records', 'other_earnings', 'other_deductions']

# ---------- db ----------
def db():
    if 'db' not in g:
        g.db = sqlite3.connect(DB); g.db.row_factory = sqlite3.Row; g.db.execute('pragma foreign_keys=on')
    return g.db
@app.teardown_appcontext
def _close(_):
    d = g.pop('db', None)
    if d: d.close()
def q(s, a=()): return db().execute(s, a).fetchall()
def q1(s, a=()):
    r = q(s, a); return r[0] if r else None
def ex(s, a=()):
    c = db().execute(s, a); db().commit(); return c.lastrowid
def _admin_pw():
    pw = os.environ.get('ADMIN_PASSWORD') or ('admin123' if MODE != 'server' else secrets.token_urlsafe(9))
    if MODE == 'server' and not os.environ.get('ADMIN_PASSWORD'): print('*** FIRST START - admin password:', pw, flush=True)
    return pw
def init_db():
    with app.app_context():
        db().executescript(open(os.path.join(os.path.dirname(__file__), 'schema.sql')).read())
        for r in ROLES: db().execute('insert or ignore into roles(name) values(?)', (r,))
        for k, v in DEFAULTS.items(): db().execute('insert or ignore into settings values(?,?)', (k, v))
        if not q1('select 1 from users'):
            db().execute('insert into users(username,full_name,password_hash,role) values(?,?,?,?)',
                         ('admin', 'Administrator', generate_password_hash(_admin_pw()), 'ADMIN'))
        for al in ('alter table contract_worker_balances add column other_earnings real default 0', 'alter table overtime_records add column work_date text', 'alter table night_work_records add column work_date text'):
            try: db().execute(al)
            except sqlite3.OperationalError: pass
        for t in UT.values():
            try: db().execute(f'alter table {t} add column uid text')
            except sqlite3.OperationalError: pass
            db().execute(f"create trigger if not exists {t}_uid after insert on {t} when new.uid is null begin update {t} set uid = lower(hex(randomblob(16))) where id = new.id; end")
            db().execute(f"update {t} set uid = lower(hex(randomblob(16))) where uid is null")
            db().execute(f'create unique index if not exists {t}_uid_ix on {t}(uid)')
        db().commit()
def S():
    s = {r['key']: r['value'] for r in q('select * from settings')}
    return {k: (s[k] if k == 'normal_end' else float(s[k])) for k in DEFAULTS}
def audit(action, entity, old='', new='', reason=''):
    ex('insert into payroll_audit_logs(at,user,action,entity,old_value,new_value,reason) values(?,?,?,?,?,?,?)',
       (dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), (me() or {}).get('username', '-'), action, entity, str(old), str(new), reason))
def fv(x):
    try: return float(str(x).replace(',', '').strip() or 0)
    except ValueError: return 0.0
def n2(x): return f'{x or 0:,.2f}'

# ---------- auth ----------
def me(): return session.get('u')
def need(*roles):
    def d(f):
        @wraps(f)
        def w(*a, **k):
            u = me()
            if not u: return redirect('/login')
            if roles and u['role'] not in roles + ('ADMIN',):
                return page('Forbidden', '<p class=err>You do not have permission for this action.</p>'), 403
            return f(*a, **k)
        return w
    return d
def back(): return redirect(request.referrer or '/')

# ---------- ui ----------
CSS = """:root{--ac:#2563eb;--bg:#f1f5f9;--bd:#e2e8f0;--tx:#0f172a;--mu:#64748b}
*{box-sizing:border-box}
body{margin:0;display:flex;min-height:100vh;background:var(--bg);color:var(--tx);font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif}
aside{width:250px;flex:none;position:sticky;top:0;height:100vh;overflow-y:auto;background:linear-gradient(180deg,#0f172a,#1e293b);color:#cbd5e1;padding:18px 12px;display:flex;flex-direction:column;gap:2px}
.brand{display:flex;align-items:center;gap:10px;color:#fff;font-weight:700;font-size:17px;padding:2px 8px 12px}
.logo{width:36px;height:36px;border-radius:10px;background:linear-gradient(135deg,#2563eb,#06b6d4);display:grid;place-items:center;font-size:18px}
.nsec{font-size:10.5px;text-transform:uppercase;letter-spacing:.09em;color:#64748b;padding:12px 12px 4px}
nav a{display:flex;align-items:center;gap:10px;color:#cbd5e1;text-decoration:none;padding:8px 12px;border-radius:8px;font-weight:500}
nav a i{font-style:normal;width:20px;text-align:center}nav a:hover{background:#1e2b45;color:#fff}nav a.on{background:var(--ac);color:#fff}
aside .box{margin-top:12px;background:#1e293b;border:1px solid #334155;border-radius:10px;padding:10px;display:flex;flex-direction:column;gap:8px}
aside label{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#94a3b8}aside select,aside button{width:100%}aside select{background:#0f172a;color:#fff;border-color:#334155}
aside .user{margin-top:auto;padding:12px 8px 0;font-size:12.5px;color:#94a3b8;border-top:1px solid #334155}aside .user a{color:#93c5fd}
main{flex:1;min-width:0;padding:22px 28px;max-width:1500px}
.top{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap;margin-bottom:16px}
h1{margin:0;font-size:24px;letter-spacing:-.01em}.sub{color:var(--mu);margin-top:2px;max-width:720px}.topr{text-align:right}.pname{font-weight:600;margin-bottom:6px}
.stepper{display:flex;gap:4px;flex-wrap:wrap;justify-content:flex-end}.stepper .s{font-size:11.5px;padding:3px 10px;border-radius:99px;background:#e2e8f0;color:#64748b;font-weight:600}
.stepper .s.done{background:#dcfce7;color:#166534}.stepper .s.cur{background:var(--ac);color:#fff}
h3{margin:0 0 10px;font-size:15px}
.card{background:#fff;border:1px solid var(--bd);border-radius:12px;padding:16px;margin:0 0 14px;box-shadow:0 1px 2px rgba(15,23,42,.04)}.card h3 small{font-weight:400;color:var(--mu)}
.g{display:flex;gap:12px;flex-wrap:wrap;align-items:flex-end}
.f{display:flex;flex-direction:column;gap:4px;font-size:12px;font-weight:600;color:#475569}.f em{font-weight:400;font-style:normal;color:var(--mu)}
input,select,button{font:inherit;padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;background:#fff;color:var(--tx)}
input:focus,select:focus{outline:2px solid #bfdbfe;border-color:var(--ac)}
button{cursor:pointer;font-weight:600}button:hover{background:#f1f5f9}button.p,aside button{background:var(--ac);border-color:var(--ac);color:#fff}button.p:hover,aside button:hover{background:#1d4ed8}
.w{overflow-x:auto;background:#fff;border:1px solid var(--bd);border-radius:12px;margin:0 0 14px}
table{border-collapse:collapse;width:100%}td,th{padding:9px 12px;border-bottom:1px solid #f1f5f9;text-align:left;white-space:nowrap}
th{background:#f8fafc;color:#475569;font-size:11.5px;text-transform:uppercase;letter-spacing:.05em}td.n,th.n{text-align:right;font-variant-numeric:tabular-nums}
tr:hover td{background:#f8fafc}tr.tot td{font-weight:700;background:#eff6ff;border-top:2px solid #bfdbfe}td input{width:96px;padding:5px 7px}
.b{display:inline-block;padding:2px 9px;border-radius:99px;font-size:11.5px;font-weight:700}
.b.ok{background:#dcfce7;color:#166534}.b.mu{background:#e2e8f0;color:#475569}.b.warn{background:#fef3c7;color:#92400e}.b.info{background:#dbeafe;color:#1e40af}.b.dark{background:#0f172a;color:#fff}.b.bad{background:#fee2e2;color:#991b1b}.alert.ok{background:#f0fdf4;border-color:#bbf7d0;color:#166534}.syncchip{display:block;margin-top:10px;background:#1e293b;border:1px solid #334155;border-radius:10px;padding:8px 10px;font-size:12.5px;color:#e2e8f0;text-decoration:none}
.alert,.err{background:#fef2f2;border:1px solid #fecaca;color:#991b1b;padding:10px 14px;border-radius:10px;margin-bottom:12px}
.muted,small{color:var(--mu)}.empty{padding:18px;text-align:center;color:var(--mu);background:#fff;border:1px dashed #cbd5e1;border-radius:12px;margin:0 0 14px}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(165px,1fr));gap:12px;margin-bottom:14px}
.kpi{background:#fff;border:1px solid var(--bd);border-radius:12px;padding:14px 16px;border-top:3px solid var(--ac)}.kpi.t{border-top-color:#14b8a6}.kpi.o{border-top-color:#f59e0b}.kpi.r{border-top-color:#ef4444}
.kpi small{font-weight:600;font-size:12px}.kpi b{display:block;font-size:22px;margin-top:4px;font-variant-numeric:tabular-nums}
.sect{font-size:13px;text-transform:uppercase;letter-spacing:.08em;color:var(--mu);margin:22px 0 8px;font-weight:700}
.grid2{display:grid;grid-template-columns:repeat(auto-fit,minmax(380px,1fr));gap:14px}
.hbars{display:flex;flex-direction:column;gap:8px}.hb{display:grid;grid-template-columns:130px 1fr 90px;gap:10px;align-items:center}
.hl{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:13px}.ht{background:#eef2f7;border-radius:99px;height:10px;overflow:hidden}.hf{height:100%;border-radius:99px}.hv{text-align:right;font-variant-numeric:tabular-nums;font-size:13px;font-weight:600}
.stack{display:flex;height:22px;border-radius:99px;overflow:hidden;background:#eef2f7;margin:8px 0}
.legend{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:12.5px;color:#475569;margin-top:6px}.legend i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:5px}.legend em{color:var(--mu);font-style:normal}
.vchart{display:flex;gap:10px;height:190px;padding-top:8px}.vc{flex:1;display:flex;flex-direction:column;height:100%}.vb{display:flex;gap:3px;align-items:flex-end;flex:1;justify-content:center;border-bottom:1px solid var(--bd)}
.vb i{display:block;width:38%;border-radius:4px 4px 0 0;min-height:2px}.vc span{font-size:11px;color:var(--mu);text-align:center;margin-top:4px}
.steps{margin:0;padding:0;list-style:none;display:grid;gap:8px}.steps li{display:flex;gap:10px;align-items:center}
.steps .ck{width:22px;height:22px;border-radius:99px;display:grid;place-items:center;background:#e2e8f0;color:#64748b;font-size:12px;font-weight:700}.steps .done .ck{background:#16a34a;color:#fff}.steps .done a{color:var(--mu);text-decoration:line-through}
a{color:var(--ac)}.pvbox{margin-top:10px;padding:8px 12px;background:#eff6ff;border-radius:8px;color:#1e40af;font-weight:600}
.loginwrap{margin:auto;padding:20px;width:100%}.logincard{max-width:380px;margin:8vh auto;background:#fff;border-radius:16px;border:1px solid var(--bd);padding:28px;box-shadow:0 10px 30px rgba(15,23,42,.08)}
.logincard .brand{color:var(--tx);padding-left:0}.logincard form{margin-top:14px}.logincard .f input,.logincard button{width:100%}
@media(max-width:900px){body{display:block}aside{width:auto;height:auto;position:static}nav{display:flex;flex-wrap:wrap}main{padding:14px}.hb{grid-template-columns:90px 1fr 80px}.grid2{grid-template-columns:1fr}}
@media print{aside,.stepper{display:none}main{padding:0}}
"""
NAV = [('Overview', [('/', 'Dashboard', '📊')]), ('Contract workers', [('/workers', 'Workers & prices', '👷'), ('/designs', 'Designs', '🪑'), ('/contract', 'Contract payroll', '🧾')]),
       ('Main employees', [('/employees', 'Employees', '👔'), ('/main', 'Main payroll', '💼')]), ('Payroll', [('/adjustments', 'Adjustments', '✏️'), ('/periods', 'Periods & approval', '📅')]),
       ('Admin', [('/settings', 'Settings', '⚙️'), ('/users', 'Users & roles', '🔑'), ('/sync', 'Sync & devices', '🔄'), ('/audit', 'Audit log', '🕵️')])]
def periods(): return q('select * from payroll_periods order by start_date, id')
def cur():
    pid = request.values.get('p')
    r = q1('select * from payroll_periods where id=?', (pid,)) if pid else None
    return r or q1('select * from payroll_periods order by start_date desc, id desc limit 1')
STAGES = ['DRAFT', 'CALCULATED', 'CHECKED', 'APPROVED', 'PAID', 'LOCKED']
HINTS = {'home': 'Overview of the selected payroll period: money in, money owed, and trends.',
         'workers': 'Register contract workers. Open a worker to set prices and see the statement and history.',
         'worker': 'Profile, worker-specific prices, price history, running account statement and payments.',
         'designs': 'Master list of designs/products. Prices are set per worker inside each worker profile.',
         'contract': 'Step 1: Add Work (price fills automatically).  Step 2: enter the amount paid.  Step 3: use the blue button at the left to move the period forward.',
         'employees': 'Register main employees with monthly gross salary and payment frequency.',
         'employee': 'Profile, salary history, payroll history and overtime history.',
         'main': 'Enter overtime, Sunday, night, transport, commission, absence and deductions. Everything else is calculated.',
         'adjustments': 'Other earnings and corrections. This is the correct way to fix a LOCKED period.',
         'periods': 'Weekly payroll periods and the approval history.',
         'settings': 'Rates used in every calculation. Only an administrator can change them.',
         'users': 'Add people and choose what each role may do.', 'sync_page': 'Work offline, then send your changes to the online server and download everyone else\'s.', 'audit_page': 'Every important change: who, what, old value, new value, when and why.'}
NUM = re.compile(r'^-?[\d,]+(\.\d+)?$')
BADGE = {'ACTIVE': 'ok', 'INACTIVE': 'mu', 'PAID': 'ok', 'PENDING': 'warn', 'DRAFT': 'mu', 'CALCULATED': 'info', 'CHECKED': 'info', 'APPROVED': 'ok', 'LOCKED': 'dark', 'Yes': 'ok', 'APPLIED': 'ok', 'DUPLICATE': 'mu', 'REJECTED': 'bad'}
def inject(html):
    tok = session.get('csrf', '')
    return re.sub(r'(<form method=post[^>]*>)', lambda m: m.group(1) + f'<input type=hidden name=_csrf value="{tok}">', html)
def fl(label, ctrl, hint=''): return f'<label class=f><span>{label}</span>{ctrl}' + (f'<em>{hint}</em>' if hint else '') + '</label>'
def ip(name, typ='text', **kw):
    return f'<input name={name} type={typ} ' + ' '.join(f'{k}="{E(v)}"' for k, v in kw.items()) + '>'
def page(title, body):
    u = me(); alerts = ''.join(f'<div class="alert {"ok" if c == "ok" else ""}">{E(m)}</div>' for c, m in get_flashed_messages(with_categories=True))
    head = f'<!doctype html><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1"><title>{E(title)} · Markid Payroll</title><style>{CSS}</style>'
    if not u:
        return inject(head + f'<div class=loginwrap><div class=logincard><div class=brand><div class=logo>🪑</div><span>Markid Payroll</span></div><p class=muted>Sign in to continue</p>{alerts}{body}</div></div>')
    P = cur(); pq = f'?p={P["id"]}' if P else ''
    nav = ''.join(f'<div class=nsec>{sec}</div>' + ''.join(f'<a href="{a}{pq}" class="{"on" if request.path == a or (a != "/" and request.path.startswith(a + "/")) else ""}"><i>{ic}</i>{l}</a>' for a, l, ic in items) for sec, items in NAV)
    opts = ''.join(f'<option value="{x["id"]}" {"selected" if P and x["id"] == P["id"] else ""}>{E(x["label"])} · {x["status"]}</option>' for x in periods())
    act = ''
    if P:
        nx = NEXT.get(P['status'])
        if nx and u['role'] in nx[1] + ('ADMIN',): act += f'<form method=post action="/period/{P["id"]}/advance"><button>{nx[2]} ▸</button></form>'
        if P['status'] in ('CALCULATED', 'CHECKED', 'APPROVED') and u['role'] in ('CHECKER', 'APPROVER', 'ADMIN'): act += f'<form method=post action="/period/{P["id"]}/reject"><button style="background:#475569;border-color:#475569">Send back</button></form>'
    box = f'<div class=box><label>Payroll period</label><form method=get action="{request.path}"><select name=p onchange="this.form.submit()">{opts}</select></form>{act}</div>' if P else '<div class=box><label>Payroll period</label><span style="font-size:13px">None yet - create one under Periods.</span></div>'
    stp = ''
    if P:
        k = STAGES.index(P['status'])
        stp = '<div class=stepper>' + ''.join(f'<span class="s {"done" if j < k else "cur" if j == k else ""}">{n.title()}</span>' for j, n in enumerate(STAGES)) + '</div>'
    top = f'<div class=top><div><h1>{E(title)}</h1><div class=sub>{E(HINTS.get(request.endpoint, ""))}</div></div><div class=topr><div class=pname>{E(P["label"]) if P else ""}</div>{stp}</div></div>'
    chip = ''
    if MODE == 'client':
        sc = sdb(); n_ = sc.execute('select count(*) from outbox').fetchone()[0]; k_ = sc.execute('select count(*) from conflicts').fetchone()[0]; sc.close()
        chip = f'<a class=syncchip href="/sync{pq}">🔄 {n_} waiting to sync' + (f' · {k_} problem(s)' if k_ else '') + f'<br><small>{E(cfg("last_sync") or "never synced")}</small></a>'
    side = f'<aside><div class=brand><div class=logo>🪑</div><span>Markid Payroll</span></div><nav>{nav}</nav>{box}{chip}<div class=user>{E(u["username"])} · {u["role"]}<br><a href="/account">Password</a> · <a href="/logout">Logout</a><div style="margin-top:6px;font-size:11px;color:#64748b">{VERSION}</div></div></aside>'
    return inject(head + f'{side}<main>{top}{alerts}{body}</main>')
def T(head, rows, total=None):
    if not rows: return '<div class=empty>No records yet.</div>'
    nc = [bool(NUM.match(str(v))) and not isinstance(v, Markup) for v in rows[0]]
    def c(x, j):
        if isinstance(x, Markup): return f'<td>{x}</td>'
        t = ('%g' % x) if isinstance(x, float) else str(x if x is not None else '')
        if t in BADGE: return f'<td><span class="b {BADGE[t]}">{E(t)}</span></td>'
        return f'<td class="{"n" if (NUM.match(t) and j < len(nc) and nc[j]) else ""}">{E(t)}</td>'
    h = ''.join(f'<th class="{"n" if (j < len(nc) and nc[j]) else ""}">{E(x)}</th>' for j, x in enumerate(head))
    b = ''.join('<tr>' + ''.join(c(x, j) for j, x in enumerate(r)) + '</tr>' for r in rows)
    if total: b += '<tr class=tot>' + ''.join(c(x, j) for j, x in enumerate(total)) + '</tr>'
    return f'<div class=w><table><tr>{h}</tr>{b}</table></div>'
def hbars(items, color='#2563eb'):
    if not items or not any(v for _, v in items): return '<div class=empty>No data for this period yet.</div>'
    m = max(v for _, v in items) or 1
    return '<div class=hbars>' + ''.join(f'<div class=hb><span class=hl title="{E(l)}">{E(l)}</span><div class=ht><div class=hf style="width:{max(v, 0) / m * 100:.1f}%;background:{color}"></div></div><span class=hv>{n2(v)}</span></div>' for l, v in items) + '</div>'
def stack(parts):
    tot = sum(v for _, v, _ in parts)
    if tot <= 0: return '<div class=empty>No payroll figures yet.</div>'
    bar = ''.join(f'<div style="width:{v / tot * 100:.2f}%;background:{c}" title="{E(l)}: {n2(v)}"></div>' for l, v, c in parts if v > 0)
    leg = ''.join(f'<span><i style="background:{c}"></i>{E(l)} <b>{n2(v)}</b> <em>{v / tot * 100:.0f}%</em></span>' for l, v, c in parts if v > 0)
    return f'<div class=stack>{bar}</div><div class=legend>{leg}</div>'
def vbars(data):
    m = max([x for _, a, b in data for x in (a, b)] or [0])
    if m <= 0: return '<div class=empty>Trend appears once periods have data.</div>'
    cols = ''.join(f'<div class=vc><div class=vb><i style="height:{a / m * 100:.1f}%;background:#2563eb" title="Contract work {n2(a)}"></i><i style="height:{b / m * 100:.1f}%;background:#14b8a6" title="Main net pay {n2(b)}"></i></div><span>{E(l)}</span></div>' for l, a, b in data)
    return f'<div class=vchart>{cols}</div><div class=legend><span><i style="background:#2563eb"></i>Contract work earned</span><span><i style="background:#14b8a6"></i>Main employees net pay</span></div>'
def hp(P): return f'<input type=hidden name=p value={P["id"]}>' if P else ''
def sel(name, rows, label='full_name', cur_id=None):
    return f'<select name={name} id={name}>' + ''.join(f'<option value={r["id"]} {"selected" if r["id"] == cur_id else ""}>{E(r[label])}</option>' for r in rows) + '</select>'
def blocked(P):
    if not P: flash('Create a payroll period first.'); return True
    if P['status'] not in ('DRAFT', 'CALCULATED'):
        flash(f'Period is {P["status"]}: editing is closed. Send it back or (admin) unlock with a reason.'); return True
    return False
def touch(P):
    if P['status'] == 'CALCULATED':
        ex("update payroll_periods set status='DRAFT' where id=?", (P['id'],)); audit('Status reset by edit', P['label'], 'CALCULATED', 'DRAFT')
def editable(P): return bool(P) and P['status'] in ('DRAFT', 'CALCULATED')

@app.route('/version')
def version(): return VERSION
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        u = q1('select * from users where username=? and active=1', (request.form['username'].strip(),))
        if u and check_password_hash(u['password_hash'], request.form['password']):
            session['u'] = {'id': u['id'], 'username': u['username'], 'role': u['role']}; audit('Login', u['username']); return redirect('/')
        flash('Invalid username or password.')
    return page('Login', '<form method=post><div class=g style="flex-direction:column;align-items:stretch">' + fl('Username', ip('username', autofocus='autofocus')) + fl('Password', ip('password', 'password')) + '<button class=p>Sign in</button></div></form>')
@app.route('/logout')
def logout(): session.clear(); return redirect('/login')

# ---------- contract logic ----------
def price_for(w, d, start):
    r = q1('select price from contract_worker_design_prices where worker_id=? and design_id=? and effective_date<=? and (end_date is null or end_date>=?) order by effective_date desc,id desc limit 1', (w, d, start, start))
    return r['price'] if r else None
def earned(w, p): return q1('select coalesce(sum(i.quantity*i.unit_price),0) t from contract_work_items i join contract_work c on c.id=i.work_id where c.worker_id=? and c.period_id=?', (w, p))['t']
def paid(w, p): return (q1('select amount from contract_worker_payments where worker_id=? and period_id=?', (w, p)) or {'amount': 0})['amount']
def other(w, p): return q1("select coalesce(sum(amount),0) t from payroll_adjustments where kind='CONTRACT' and person_id=? and period_id=?", (w, p))['t']
def bal(w, pid):
    prev = q1('select opening_balance o from contract_workers where id=?', (w,))['o'] or 0
    for P in periods():
        e, o, pd = earned(w, P['id']), other(w, P['id']), paid(w, P['id']); pay = prev + e + o
        r = dict(period=P, prev=prev, earned=e, other=o, payable=pay, paid=pd, close=pay - pd)
        if P['id'] == pid: return r
        prev = r['close']
    return dict(prev=prev, earned=0, other=0, payable=prev, paid=0, close=prev)
def contract_rows(P):
    out = []
    for w in q('select * from contract_workers order by full_name'):
        b = bal(w['id'], P['id'])
        if w['status'] == 'ACTIVE' or b['earned'] or b['other'] or b['close']: out.append((w, b))
    return out

# ---------- main logic ----------
def minutes(t): h, m = str(t).split(':')[:2]; return int(h) * 60 + int(m)
def gross_at(eid, start):
    r = q1('select new_gross g from main_employee_salary_history where employee_id=? and effective_date<=? order by effective_date desc,id desc limit 1', (eid, start))
    return r['g'] if r else q1('select monthly_gross g from main_employees where id=?', (eid,))['g']
def is_due(freq, wk): return freq == 'WEEKLY' or (freq == 'BIWEEKLY' and wk % 2 == 0) or (freq == 'MONTHLY' and wk % 4 == 0)
def inp(pid, e):
    a = (pid, e); one = lambda t: q1(f'select * from {t} where period_id=? and employee_id=?', a)
    hrs = lambda t, c='': q1(f'select coalesce(sum(hours),0) h from {t} where period_id=? and employee_id=?{c}', a)['h']
    tr, co, ab, oe, od = (one(t) for t in ('transport_records', 'commission_records', 'absence_records', 'other_earnings', 'other_deductions'))
    v = lambda r, k: (r[k] if r and r[k] is not None else 0)
    t = lambda r, k: (r[k] if r and r[k] else '')
    return dict(oh=hrs('overtime_records', " and ot_type='NORMAL'"), sun=hrs('overtime_records', " and ot_type='SUNDAY'"), night=hrs('night_work_records'),
                tr=v(tr, 'transport'), ex=v(tr, 'expense'), al=v(tr, 'other_allowance'), com=v(co, 'amount'), cd=t(co, 'description'), co=t(co, 'customer_order'), rem=t(co, 'remark'),
                oe=v(oe, 'amount'), absd=v(ab, 'absent_days'), abs=v(ab, 'deduction_amount'), od=v(od, 'amount'))
def adj_main(e, p):
    r = q1("select coalesce(sum(case when amount>0 then amount end),0) a, coalesce(sum(case when amount<0 then -amount end),0) b from payroll_adjustments where kind='MAIN' and person_id=? and period_id=?", (e, p)); return r['a'], r['b']
def calc(e, P, s):
    i = inp(P['id'], e['id']); gross = gross_at(e['id'], P['start_date']); fr = e['payment_frequency']; due = is_due(fr, P['week_number']); rate = gross / s['hours']
    basic = gross / DIV[fr] if due else 0; otp, sp, np_ = i['oh'] * rate * s['ot_mult'], i['sun'] * rate * s['sun_mult'], i['night'] * rate * s['night_mult']
    tr = i['tr'] + i['ex'] + i['al']
    pl, mi = adj_main(e['id'], P['id']); ab = i['abs'] or (gross / s['absence_divisor'] * i['absd'] if s['absence_divisor'] > 0 else 0)
    oe, od = i['oe'] + pl, i['od'] + mi
    net = basic + tr + otp + sp + np_ + i['com'] + oe - ab - od if due else 0
    return dict(employee_id=e['id'], name=e['full_name'], dept=e['department'], monthly_gross=gross, frequency=fr, due=int(due), basic=round(basic, 2), transport=round(tr, 2),
                ot_hours=i['oh'], ot_pay=round(otp, 2), sun_hours=i['sun'], sun_pay=round(sp, 2), night_hours=i['night'], night_pay=round(np_, 2), commission=i['com'],
                other_earnings=round(oe, 2), absence=round(ab, 2), other_deductions=round(od, 2), net_pay=round(net, 2), remark=i['rem'], payment_status='PENDING')
def main_rows(P):
    if P['status'] == 'DRAFT':
        s = S(); return [calc(e, P, s) for e in q("select * from main_employees where status='ACTIVE' order by full_name")]
    return [dict(r) for r in q('select r.*, e.full_name name, e.department dept from main_employee_payroll r join main_employees e on e.id=r.employee_id where r.period_id=? order by e.full_name', (P['id'],))]
def snapshot(P):
    s = S(); db().execute('delete from main_employee_payroll where period_id=?', (P['id'],))
    for e in q("select * from main_employees where status='ACTIVE'"):
        c = calc(e, P, s); cols = [k for k in c if k not in ('name', 'dept')]
        db().execute(f'insert into main_employee_payroll(period_id,{",".join(cols)}) values(?{",?" * len(cols)})', (P['id'], *[c[k] for k in cols]))
    db().execute('delete from contract_worker_balances where period_id=?', (P['id'],))
    for w, b in contract_rows(P):
        db().execute('insert into contract_worker_balances(period_id,worker_id,previous_balance,work_earned,other_earnings,total_payable,amount_paid,closing_balance) values(?,?,?,?,?,?,?,?)',
                     (P['id'], w['id'], b['prev'], b['earned'], b['other'], b['payable'], b['paid'], b['close']))
    db().commit()

# ---------- periods / workflow ----------
@app.route('/periods', methods=['GET', 'POST'])
@need()
def periods_page():
    if request.method == 'POST':
        if me()['role'] not in ('ACCOUNTANT', 'ADMIN'): return page('Forbidden', ''), 403
        f = request.form; n = int(fv(f['week']))
        pid = ex('insert into payroll_periods(label,start_date,end_date,week_number,calendar) values(?,?,?,?,?)',
                 (f['label'] or f'Markid Weekly Payroll Week {n}', f['start'], f['end'], n, f['cal'] or 'E.C.'))
        audit('Period created', f['label'], '', f['start']); return redirect(f'/periods?p={pid}')
    rows = [(x['id'], x['label'], x['week_number'], x['calendar'], x['start_date'], x['end_date'], x['status']) for x in periods()]
    form = ('<form method=post class=card><h3>New weekly payroll period</h3><div class=g>' + fl('Label', ip('label', size='36', placeholder='e.g. Markid Weekly Payroll 23/12/2018 E.C.')) + fl('Week number', ip('week', 'number', required='required', style='width:110px'), 'drives biweekly/monthly due') + fl('Start date', ip('start', 'date', required='required')) + fl('End date', ip('end', 'date', required='required')) + fl('Calendar', ip('cal', value='E.C.', style='width:80px')) + '<button class=p>Create period</button></div></form>')
    ap = T(['Period', 'Action', 'By', 'At', 'Note'], [(r['period_id'], r['from_status'] + ' → ' + r['to_status'], r['user'], r['at'], r['note']) for r in q('select * from payroll_approvals order by id desc limit 40')])
    return page('Payroll periods', form + T(['ID', 'Label', 'Week', 'Calendar', 'Start', 'End', 'Status'], rows) + '<h3>Approval history</h3>' + ap)
@app.route('/period/<int:pid>/advance', methods=['POST'])
@need()
def advance(pid):
    P = q1('select * from payroll_periods where id=?', (pid,)); nx = NEXT.get(P['status'])
    if not nx or me()['role'] not in nx[1] + ('ADMIN',): flash('Not allowed.'); return back()
    if nx[0] == 'CALCULATED': snapshot(P)
    if nx[0] == 'PAID': ex("update main_employee_payroll set payment_status='PAID' where period_id=?", (pid,))
    ex('update payroll_periods set status=? where id=?', (nx[0], pid))
    ex('insert into payroll_approvals(period_id,from_status,to_status,user,at,note) values(?,?,?,?,?,?)', (pid, P['status'], nx[0], me()['username'], dt.datetime.now().strftime('%Y-%m-%d %H:%M'), ''))
    audit('Status', P['label'], P['status'], nx[0]); return back()
@app.route('/period/<int:pid>/reject', methods=['POST'])
@need('CHECKER', 'APPROVER')
def reject(pid):
    P = q1('select * from payroll_periods where id=?', (pid,))
    if P['status'] in ('CALCULATED', 'CHECKED', 'APPROVED'):
        ex("update payroll_periods set status='DRAFT' where id=?", (pid,)); audit('Sent back', P['label'], P['status'], 'DRAFT')
        ex('insert into payroll_approvals(period_id,from_status,to_status,user,at,note) values(?,?,?,?,?,?)', (pid, P['status'], 'DRAFT', me()['username'], dt.datetime.now().strftime('%Y-%m-%d %H:%M'), 'sent back'))
    return back()
@app.route('/period/<int:pid>/unlock', methods=['POST'])
@need('ADMIN')
def unlock(pid):
    r = request.form.get('reason', '').strip()
    if not r: flash('A reason is required to unlock a locked period.'); return back()
    P = q1('select * from payroll_periods where id=?', (pid,)); ex("update payroll_periods set status='DRAFT' where id=?", (pid,))
    ex('insert into payroll_approvals(period_id,from_status,to_status,user,at,note) values(?,?,?,?,?,?)', (pid, P['status'], 'DRAFT', me()['username'], dt.datetime.now().strftime('%Y-%m-%d %H:%M'), 'correction: ' + r))
    audit('Correction unlock', P['label'], P['status'], 'DRAFT', r); return back()

# ---------- dashboard ----------
def checklist():
    st = [('Create a payroll period (weekly)', '/periods', q1('select 1 x from payroll_periods')), ('Add designs / products', '/designs', q1('select 1 x from designs')),
          ('Add contract workers', '/workers', q1('select 1 x from contract_workers')), ('Open a worker and set their price for each design', '/workers', q1('select 1 x from contract_worker_design_prices')),
          ('Add main employees', '/employees', q1('select 1 x from main_employees')), ("Enter this week's contract work (Contract Payroll)", '/contract', q1('select 1 x from contract_work_items')),
          ('Enter overtime / transport for main employees (Main Payroll)', '/main', q1('select 1 x from overtime_records union all select 1 from transport_records limit 1'))]
    done = sum(1 for _, _, d in st if d)
    lis = ''.join(f'<li class="{"done" if d else ""}"><span class=ck>{"✓" if d else i}</span><a href="{u}">{E(t)}</a></li>' for i, (t, u, d) in enumerate(st, 1))
    return f'<div class=card><h3>Getting started <small>{done} of {len(st)} steps done</small></h3><ol class=steps>{lis}</ol></div>' if done < len(st) else ''
def trend(n=8):
    out = []
    for X in periods()[-n:]:
        cw = q1('select coalesce(sum(i.quantity*i.unit_price),0) t from contract_work_items i join contract_work c on c.id=i.work_id where c.period_id=?', (X['id'],))['t']
        out.append((f'Wk {X["week_number"]}', cw, sum(r['net_pay'] for r in main_rows(X) if r['due'])))
    return out
@app.route('/')
@need()
def home():
    P = cur(); cl = checklist()
    if not P: return page('Dashboard', cl + '<div class=empty>No payroll period yet. Create one in <a href="/periods">Periods</a> to start.</div>')
    cr = contract_rows(P); mr = [r for r in main_rows(P) if r['due']]; sm = lambda a, k: sum(x[k] for x in a); bs = [b for _, b in cr]
    k = lambda l, v, c='': f'<div class="kpi {c}"><small>{l}</small><b>{v}</b></div>'
    ck = (k('Active contract workers', q1("select count(*) c from contract_workers where status='ACTIVE'")['c']) + k("Current week's work", n2(sm(bs, 'earned'))) + k('Other earnings', n2(sm(bs, 'other'))) + k('Previous unpaid', n2(sm(bs, 'prev')), 'o')
          + k('Total payable', n2(sm(bs, 'payable'))) + k('Total paid', n2(sm(bs, 'paid')), 't') + k('Total outstanding', n2(sm(bs, 'close')), 'r'))
    mk = (k('Main employees', q1("select count(*) c from main_employees where status='ACTIVE'")['c']) + k('Paid this period', len(mr)) + k('Monthly gross (all)', n2(q1("select coalesce(sum(monthly_gross),0) s from main_employees where status='ACTIVE'")['s']))
          + k('Basic salary', n2(sm(mr, 'basic'))) + k('Overtime', n2(sm(mr, 'ot_pay') + sm(mr, 'sun_pay')), 'o') + k('Night work', n2(sm(mr, 'night_pay'))) + k('Transport', n2(sm(mr, 'transport')))
          + k('Commission', n2(sm(mr, 'commission'))) + k('Deductions', n2(sm(mr, 'absence') + sm(mr, 'other_deductions')), 'r') + k('Total net pay', n2(sm(mr, 'net_pay')), 't'))
    parts = [('Basic', sm(mr, 'basic'), '#2563eb'), ('Overtime', sm(mr, 'ot_pay') + sm(mr, 'sun_pay'), '#f59e0b'), ('Night work', sm(mr, 'night_pay'), '#7c3aed'), ('Transport', sm(mr, 'transport'), '#14b8a6'), ('Commission', sm(mr, 'commission'), '#ec4899'), ('Other earnings', sm(mr, 'other_earnings'), '#64748b')]
    top = [(r['n'], r['v']) for r in q('select d.name n, sum(i.quantity*i.unit_price) v from contract_work_items i join contract_work c on c.id=i.work_id join designs d on d.id=i.design_id where c.period_id=? group by d.id order by v desc limit 6', (P['id'],))]
    card = lambda t, inner: f'<div class=card><h3>{t}</h3>{inner}</div>'
    t1 = T(['Worker', 'Current work', 'Other', 'Previous', 'Payable', 'Paid', 'Outstanding'], [(w['full_name'], n2(b['earned']), n2(b['other']), n2(b['prev']), n2(b['payable']), n2(b['paid']), n2(b['close'])) for w, b in cr])
    t2 = T(['Employee', 'Monthly gross', 'Frequency', 'Basic', 'OT', 'Night', 'Transport', 'Commission', 'Deductions', 'Net', 'Payment'],
           [(r['name'], n2(r['monthly_gross']), r['frequency'], n2(r['basic']), n2(r['ot_pay'] + r['sun_pay']), n2(r['night_pay']), n2(r['transport']), n2(r['commission']), n2(r['absence'] + r['other_deductions']), n2(r['net_pay']), r['payment_status']) for r in mr])
    note = ''
    if P['status'] == 'LOCKED':
        note = '<div class=card><b>This period is LOCKED.</b> Use <a href="/adjustments">Adjustments</a> to correct it.'
        if me()['role'] == 'ADMIN': note += f'<form method=post action="/period/{P["id"]}/unlock" class=g style="margin-top:8px">' + fl('Admin: reason to unlock for correction', ip('reason', size='40')) + '<button>Unlock</button></form>'
        note += '</div>'
    return page('Dashboard', cl + note + '<div class=sect>Contract workers</div><div class=kpis>' + ck + '</div><div class=grid2>' + card('Work earned by worker', hbars([(w['full_name'], b['earned'] + b['other']) for w, b in cr])) + card('Outstanding balance by worker', hbars([(w['full_name'], b['close']) for w, b in cr], '#ef4444')) + '</div>'
                + card('Top designs this period (value)', hbars(top, '#f59e0b')) + t1 + '<div class=sect>Main employees</div><div class=kpis>' + mk + '</div><div class=grid2>' + card('Net pay by employee', hbars([(r['name'], r['net_pay']) for r in mr], '#14b8a6'))
                + card('Where the payroll money goes', stack(parts)) + '</div>' + card('Trend: last 8 periods', vbars(trend())) + t2)

# ---------- designs ----------
@app.route('/designs', methods=['GET', 'POST'])
@need()
def designs():
    if request.method == 'POST':
        if me()['role'] not in ('ACCOUNTANT', 'ADMIN'): return page('Forbidden', ''), 403
        f = request.form
        try:
            did = ex('insert into designs(name,category,description,notes) values(?,?,?,?)', (f['name'].strip(), f['cat'], f['desc'], f['notes']))
            ex('insert into products(design_id,name,notes) values(?,?,?)', (did, f['name'].strip(), f['notes'])); audit('Design added', f['name'])
        except sqlite3.IntegrityError: flash('Design name already exists.')
        return back()
    rows = [(d['id'], d['name'], d['category'], d['description'], 'Active' if d['active'] else 'Inactive', Markup(f'<form method=post action="/designs/{d["id"]}/toggle"><button>Toggle</button></form>')) for d in q('select * from designs order by name')]
    form = '<form method=post class=card><h3>Add a design / product</h3><div class=g>' + fl('Design name', ip('name', required='required')) + fl('Category', ip('cat')) + fl('Description', ip('desc')) + fl('Notes', ip('notes')) + '<button class=p>Add design</button></div></form>'
    return page('Designs & Products', form + T(['ID', 'Design', 'Category', 'Description', 'Status', ''], rows))
@app.route('/designs/<int:i>/toggle', methods=['POST'])
@need('ACCOUNTANT')
def design_toggle(i): ex('update designs set active=1-active where id=?', (i,)); audit('Design toggled', i); return back()

# ---------- contract workers ----------
@app.route('/workers', methods=['GET', 'POST'])
@need()
def workers():
    if request.method == 'POST':
        if me()['role'] not in ('ACCOUNTANT', 'ADMIN'): return page('Forbidden', ''), 403
        f = request.form; wid = ex('insert into contract_workers(full_name,phone,start_date,notes,opening_balance) values(?,?,?,?,?)', (f['name'].strip(), f['phone'], f['start'], f['notes'], fv(f['open'])))
        audit('Worker added', f['name']); return redirect(f'/workers/{wid}')
    rows = [(w['id'], w['full_name'], w['phone'], w['start_date'], w['status'], Markup(f'<a href="/workers/{w["id"]}">Open</a>')) for w in q('select * from contract_workers order by full_name')]
    form = '<form method=post class=card><h3>Add a contract worker</h3><div class=g>' + fl('Full name', ip('name', required='required')) + fl('Phone', ip('phone')) + fl('Start date', ip('start', 'date')) + fl('Opening unpaid balance', ip('open', 'number', step='any', value='0'), 'only if already owed money') + fl('Notes', ip('notes')) + '<button class=p>Add worker</button></div></form>'
    return page('Contract Workers', form + T(['ID', 'Name', 'Phone', 'Start', 'Status', ''], rows))
@app.route('/workers/<int:i>', methods=['GET', 'POST'])
@need()
def worker(i):
    w = q1('select * from contract_workers where id=?', (i,))
    if request.method == 'POST':
        if me()['role'] not in ('ACCOUNTANT', 'ADMIN'): return page('Forbidden', ''), 403
        f = request.form
        for fld, col in (('name', 'full_name'), ('phone', 'phone'), ('start', 'start_date'), ('notes', 'notes')):
            if f[fld] != (w[col] or ''): audit(f'Worker {col}', w['full_name'], w[col], f[fld])
        ex('update contract_workers set full_name=?,phone=?,start_date=?,notes=?,status=? where id=?', (f['name'], f['phone'], f['start'], f['notes'], f['status'], i)); return back()
    P = cur(); ds = q('select * from designs where active=1 order by name')
    prices = q('select p.*, d.name dn from contract_worker_design_prices p join designs d on d.id=p.design_id where worker_id=? order by d.name, effective_date desc', (i,))
    hist = q('select h.*, d.name dn from contract_worker_design_price_history h join designs d on d.id=h.design_id where worker_id=? order by h.id desc', (i,))
    prof = ('<form method=post class=card><h3>Profile</h3><div class=g>' + fl('Full name', ip('name', value=w['full_name'])) + fl('Phone', ip('phone', value=w['phone'] or '')) + fl('Start date', ip('start', 'date', value=w['start_date'] or '')) + fl('Notes', ip('notes', value=w['notes'] or '')) + fl('Status', f'<select name=status><option {"selected" if w["status"] == "ACTIVE" else ""}>ACTIVE</option><option {"selected" if w["status"] != "ACTIVE" else ""}>INACTIVE</option></select>') + '<button class=p>Save profile</button></div></form>')
    pf = (f'<form method=post action="/workers/{i}/price" class=card><h3>Set worker price</h3><div class=g>' + fl('Design', sel('design', ds, 'name')) + fl('Price per unit (birr)', ip('price', 'number', step='any', required='required')) + fl('Effective from', ip('eff', 'date', required='required')) + fl('Notes', ip('notes')) + '<button class=p>Save price</button></div><small>A new price automatically closes the previous one. Old payroll keeps the price it was calculated with.</small></form>')
    pt = T(['Design', 'Price per unit', 'Effective', 'Ends'], [(p['dn'], n2(p['price']), p['effective_date'], p['end_date'] or 'current') for p in prices])
    ht = T(['Design', 'Old', 'New', 'Effective', 'By', 'When'], [(h['dn'], n2(h['old_price']), n2(h['new_price']), h['effective_date'], h['changed_by'], h['changed_at']) for h in hist])
    st, te, tp = [], 0, 0
    for X in periods():
        b = bal(i, X['id']); te += b['earned'] + b['other']; tp += b['paid']
        if b['earned'] or b['other'] or b['paid'] or b['prev']: st.append((X['label'], n2(b['prev']), n2(b['earned']), n2(b['other']), n2(b['payable']), n2(b['paid']), n2(b['close'])))
    out = b['close'] if periods() else (w['opening_balance'] or 0)
    stt = T(['Period', 'Opening balance', '+ Work earned', '+ Other earnings', '= Payable', '− Paid', 'Closing balance'], st)
    wh = T(['Period', 'Design', 'Qty', 'Price', 'Total', 'Customer', 'Remark'], [(r['label'], r['dn'], r['quantity'], n2(r['unit_price']), n2(r['quantity'] * r['unit_price']), r['cn'] or '', r['remark'] or '') for r in q(
        'select pp.label, d.name dn, i.quantity, i.unit_price, c.name cn, i.remark from contract_work_items i join contract_work cw on cw.id=i.work_id join payroll_periods pp on pp.id=cw.period_id join designs d on d.id=i.design_id left join customers c on c.id=i.customer_id where cw.worker_id=? order by pp.start_date, i.id', (i,))])
    pay = T(['Period', 'Paid', 'By', 'At'], [(r['label'], n2(r['amount']), r['paid_by'], r['paid_at']) for r in q('select p.*, pp.label from contract_worker_payments p join payroll_periods pp on pp.id=p.period_id where worker_id=? order by pp.start_date', (i,))])
    summ = f'<div class=card><div class=k><small>Total earned</small><b>{n2(te)}</b></div><div class=k><small>Total paid</small><b>{n2(tp)}</b></div><div class=k><small>Total unpaid</small><b>{n2(out)}</b></div></div>'
    return page(f'Worker · {w["full_name"]}', f'{prof}{summ}<h3>Worker design prices</h3>{pf}{pt}<h3>Price history</h3>{ht}<h3>Account statement</h3>{stt}<p><a href="/export/statement/{i}">⬇ Statement as Excel</a></p><h3>Work history</h3>{wh}<h3>Payment history</h3>{pay}')
@app.route('/workers/<int:i>/price', methods=['POST'])
@need('ACCOUNTANT')
def set_price(i):
    f = request.form; d = int(f['design']); p = fv(f['price']); eff = f['eff']
    if p <= 0: flash('Price must be positive.'); return back()
    old = q1('select * from contract_worker_design_prices where worker_id=? and design_id=? and end_date is null order by effective_date desc limit 1', (i, d))
    if old and old['effective_date'] >= eff: flash('Effective date must be after the current price start date.'); return back()
    if old: ex('update contract_worker_design_prices set end_date=? where id=?', ((dt.date.fromisoformat(eff) - dt.timedelta(days=1)).isoformat(), old['id']))
    ex('insert into contract_worker_design_prices(worker_id,design_id,price,effective_date,created_by,notes) values(?,?,?,?,?,?)', (i, d, p, eff, me()['username'], f['notes']))
    ex('insert into contract_worker_design_price_history(worker_id,design_id,old_price,new_price,effective_date,changed_by,changed_at,notes) values(?,?,?,?,?,?,?,?)',
       (i, d, old['price'] if old else None, p, eff, me()['username'], dt.datetime.now().strftime('%Y-%m-%d %H:%M'), f['notes']))
    audit('Price', f'{q1("select full_name n from contract_workers where id=?", (i,))["n"]} / {q1("select name n from designs where id=?", (d,))["n"]}', old['price'] if old else '-', p); return back()

# ---------- contract payroll ----------
@app.route('/contract')
@need()
def contract():
    P = cur()
    if not P: return page('Contract Payroll', '<p>Create a payroll period first.</p>')
    ws = q("select * from contract_workers where status='ACTIVE' order by full_name"); ds = q('select * from designs where active=1 order by name')
    pr = {f'{w["id"]}_{d["id"]}': price_for(w['id'], d['id'], P['start_date']) for w in ws for d in ds}; pr = {k: v for k, v in pr.items() if v is not None}
    ed = editable(P)
    form = (f'<form method=post action="/contract/add" class=card>{hp(P)}<h3>Add work</h3><div class=g>' + fl('Worker', sel('w', ws)) + fl('Design', sel('d', ds, 'name')) + fl('Quantity', ip('qty', 'number', id='qq', step='any', required='required', style='width:110px')) + fl('Customer (optional)', ip('cust')) + fl('Remark', ip('remark', size='28')) + '<button class=p>Add Work</button></div><div class=pvbox id=pvx>Price fills automatically from the worker price list.</div></form>'
            + f'<script>const PR={json.dumps(pr)};function pv(){{const k=w.value+"_"+d.value,p=PR[k],q=+qq.value||0;pvx.textContent=p==null?"No price set for this worker/design - set it in the worker profile":"Price "+p+" x "+q+" = "+(p*q)}};w.onchange=d.onchange=qq.oninput=pv;pv()</script>') if ed else f'<p class=err>Period is {P["status"]} - read only.</p>'
    items = q('select i.id, w.full_name wn, d.name dn, i.quantity, i.unit_price, c.name cn, i.remark from contract_work_items i join contract_work cw on cw.id=i.work_id join contract_workers w on w.id=cw.worker_id join designs d on d.id=i.design_id left join customers c on c.id=i.customer_id where cw.period_id=? order by w.full_name,i.id', (P['id'],))
    rows = [(r['wn'], r['dn'], r['quantity'], n2(r['unit_price']), n2(r['quantity'] * r['unit_price']), r['cn'] or '', r['remark'] or '', Markup(f'<form method=post action="/contract/del/{r["id"]}">{hp(P)}<button>×</button></form>') if ed else '') for r in items]
    cr = contract_rows(P)
    srows = [(w['full_name'], n2(b['prev']), n2(b['earned']), n2(b['other']), n2(b['payable']), Markup(f'<input name=pay_{w["id"]} type=number step=any value="{b["paid"]}" {"" if ed else "disabled"}>'), n2(b['close'])) for w, b in cr]
    sm = lambda k: n2(sum(b[k] for _, b in cr)); tot = ('TOTAL', sm('prev'), sm('earned'), sm('other'), sm('payable'), sm('paid'), sm('close'))
    sform = f'<form method=post action="/contract/pay">{hp(P)}{T(["Worker", "Previous", "Current work", "Other / adjust.", "Payable", "Paid", "Balance"], srows, tot)}{"<button class=p>Save payments</button>" if ed else ""}</form>'
    return page(f'Contract Payroll · {P["label"]}', f'{form}<h3>Work entries</h3>{T(["Worker", "Design", "Qty", "Price", "Total", "Customer", "Remark", ""], rows)}<h3>Summary &amp; payments</h3>{sform}<p><a href="/export/contract?p={P["id"]}">⬇ Excel report</a></p>')
@app.route('/contract/add', methods=['POST'])
@need('ACCOUNTANT')
def contract_add():
    P = cur(); f = request.form
    if blocked(P): return back()
    w, d, qty = int(f['w']), int(f['d']), fv(f['qty']); pr = price_for(w, d, P['start_date'])
    if pr is None: flash("No price set for this worker and design for this period. Add it in the worker's profile first."); return back()
    if qty <= 0: flash('Quantity must be positive.'); return back()
    cid = None
    if f['cust'].strip():
        db().execute('insert or ignore into customers(name) values(?)', (f['cust'].strip(),)); cid = q1('select id from customers where name=?', (f['cust'].strip(),))['id']
    db().execute('insert or ignore into contract_work(period_id,worker_id) values(?,?)', (P['id'], w)); wid = q1('select id from contract_work where period_id=? and worker_id=?', (P['id'], w))['id']
    ex('insert into contract_work_items(work_id,design_id,quantity,unit_price,customer_id,remark) values(?,?,?,?,?,?)', (wid, d, qty, pr, cid, f['remark'])); touch(P); return back()
@app.route('/contract/del/<int:i>', methods=['POST'])
@need('ACCOUNTANT')
def contract_del(i):
    P = cur()
    if blocked(P): return back()
    r = q1('select i.*, w.full_name wn, d.name dn from contract_work_items i join contract_work cw on cw.id=i.work_id join contract_workers w on w.id=cw.worker_id join designs d on d.id=i.design_id where i.id=?', (i,))
    audit('Work entry deleted', f'{r["wn"]} / {r["dn"]}', f'{r["quantity"]} x {r["unit_price"]}', ''); ex('delete from contract_work_items where id=?', (i,)); touch(P); return back()
@app.route('/contract/pay', methods=['POST'])
@need('ACCOUNTANT')
def contract_pay():
    P = cur()
    if blocked(P): return back()
    for k, v in request.form.items():
        if not k.startswith('pay_'): continue
        w = int(k[4:]); new = fv(v); old = paid(w, P['id'])
        if new != old:
            db().execute('insert into contract_worker_payments(period_id,worker_id,amount,paid_by,paid_at) values(?,?,?,?,?) on conflict(period_id,worker_id) do update set amount=excluded.amount,paid_by=excluded.paid_by,paid_at=excluded.paid_at',
                         (P['id'], w, new, me()['username'], dt.datetime.now().strftime('%Y-%m-%d %H:%M')))
            audit('Contract payment', f'{q1("select full_name n from contract_workers where id=?", (w,))["n"]} / {P["label"]}', old, new)
    db().commit(); touch(P); return back()

# ---------- main employees ----------
@app.route('/employees', methods=['GET', 'POST'])
@need()
def employees():
    if request.method == 'POST':
        if me()['role'] not in ('ACCOUNTANT', 'ADMIN'): return page('Forbidden', ''), 403
        f = request.form; g_ = fv(f['gross'])
        eid = ex('insert into main_employees(full_name,department,position,monthly_gross,payment_frequency,start_date,notes) values(?,?,?,?,?,?,?)', (f['name'].strip(), f['dept'], f['pos'], g_, f['freq'], f['start'], f['notes']))
        ex('insert into main_employee_salary_history(employee_id,old_gross,new_gross,effective_date,reason,changed_by) values(?,?,?,?,?,?)', (eid, 0, g_, '0000-01-01', 'Initial salary', me()['username']))
        audit('Employee added', f['name'], '', g_); return redirect(f'/employees/{eid}')
    rows = [(e['id'], e['full_name'], e['department'], e['position'], n2(e['monthly_gross']), e['payment_frequency'], e['status'], Markup(f'<a href="/employees/{e["id"]}">Open</a>')) for e in q('select * from main_employees order by full_name')]
    form = ('<form method=post class=card><h3>Add a main employee</h3><div class=g>' + fl('Full name', ip('name', required='required')) + fl('Department', ip('dept')) + fl('Position', ip('pos')) + fl('Monthly gross salary', ip('gross', 'number', step='any', required='required')) + fl('Payment frequency', '<select name=freq><option>WEEKLY</option><option>BIWEEKLY</option><option>MONTHLY</option></select>') + fl('Start date', ip('start', 'date')) + fl('Notes', ip('notes')) + '<button class=p>Add employee</button></div></form>')
    return page('Main Employees', form + T(['ID', 'Name', 'Department', 'Position', 'Monthly gross', 'Frequency', 'Status', ''], rows))
@app.route('/employees/<int:i>', methods=['GET', 'POST'])
@need()
def employee(i):
    e = q1('select * from main_employees where id=?', (i,))
    if request.method == 'POST':
        if me()['role'] not in ('ACCOUNTANT', 'ADMIN'): return page('Forbidden', ''), 403
        f = request.form
        for fld, col in (('name', 'full_name'), ('dept', 'department'), ('pos', 'position'), ('freq', 'payment_frequency'), ('status', 'status')):
            if f[fld] != (e[col] or ''): audit(f'Employee {col}', e['full_name'], e[col], f[fld])
        ex('update main_employees set full_name=?,department=?,position=?,payment_frequency=?,status=?,notes=? where id=?', (f['name'], f['dept'], f['pos'], f['freq'], f['status'], f['notes'], i)); return back()
    o = lambda cur_, vals: ''.join(f'<option {"selected" if v == cur_ else ""}>{v}</option>' for v in vals)
    prof = ('<form method=post class=card><h3>Profile</h3><div class=g>' + fl('Full name', ip('name', value=e['full_name'])) + fl('Department', ip('dept', value=e['department'] or '')) + fl('Position', ip('pos', value=e['position'] or '')) + fl('Payment frequency', f'<select name=freq>{o(e["payment_frequency"], DIV)}</select>') + fl('Status', f'<select name=status>{o(e["status"], ["ACTIVE", "INACTIVE"])}</select>') + fl('Notes', ip('notes', value=e['notes'] or '')) + '<button class=p>Save profile</button></div></form>')
    sal = (f'<form method=post action="/employees/{i}/salary" class=card><h3>Change monthly gross <small>(current {n2(e["monthly_gross"])})</small></h3><div class=g>' + fl('New monthly gross', ip('gross', 'number', step='any', required='required')) + fl('Effective from', ip('eff', 'date', required='required')) + fl('Reason', ip('reason', required='required')) + '<button class=p>Change salary</button></div></form>')
    sh = T(['Effective', 'Old', 'New', 'Reason', 'By'], [(r['effective_date'], n2(r['old_gross']), n2(r['new_gross']), r['reason'], r['changed_by']) for r in q('select * from main_employee_salary_history where employee_id=? order by effective_date desc,id desc', (i,))])
    ph = T(['Period', 'Gross', 'Freq', 'Basic', 'OT+Sunday', 'Night', 'Transport', 'Commission', 'Deductions', 'Net', 'Status'],
           [(r['label'], n2(r['monthly_gross']), r['frequency'], n2(r['basic']), n2(r['ot_pay'] + r['sun_pay']), n2(r['night_pay']), n2(r['transport']), n2(r['commission']), n2(r['absence'] + r['other_deductions']), n2(r['net_pay']), r['payment_status'])
            for r in q('select r.*, p.label from main_employee_payroll r join payroll_periods p on p.id=r.period_id where employee_id=? and due=1 order by p.start_date', (i,))])
    oh = T(['Period', 'Date', 'Type', 'Normal end', 'Finished at', 'Hours'], [(r['label'], r['d'] or '', r['t'], r['ne'] or '', r['ae'] or '', r['h']) for r in q("select p.label label, o.work_date d, o.ot_type t, o.normal_end ne, o.actual_end ae, o.hours h, p.start_date sd from overtime_records o join payroll_periods p on p.id=o.period_id where o.employee_id=? union all select p.label, o.work_date, 'NIGHT', '', '', o.hours, p.start_date from night_work_records o join payroll_periods p on p.id=o.period_id where o.employee_id=? order by sd, d", (i, i))])
    return page(f'Employee · {e["full_name"]}', f'{prof}{sal}<h3>Salary history</h3>{sh}<h3>Payroll &amp; payment history</h3>{ph}<h3>Overtime history</h3>{oh}')
@app.route('/employees/<int:i>/salary', methods=['POST'])
@need('ACCOUNTANT')
def salary(i):
    f = request.form; e = q1('select * from main_employees where id=?', (i,)); new = fv(f['gross'])
    if new <= 0: flash('Gross must be positive.'); return back()
    ex('insert into main_employee_salary_history(employee_id,old_gross,new_gross,effective_date,reason,changed_by) values(?,?,?,?,?,?)', (i, e['monthly_gross'], new, f['eff'], f['reason'], me()['username']))
    latest = q1('select new_gross g from main_employee_salary_history where employee_id=? order by effective_date desc,id desc limit 1', (i,))['g']
    ex('update main_employees set monthly_gross=? where id=?', (latest, i)); audit('Employee gross salary', e['full_name'], e['monthly_gross'], new, f['reason']); return back()

# ---------- main payroll ----------
FIELDS = [('tr', 'Transport', 'number'), ('ex', 'Expense', 'number'),
          ('al', 'Other allow.', 'number'), ('com', 'Commission', 'number'), ('cd', 'Comm. description', 'text'), ('co', 'Customer/order', 'text'), ('oe', 'Other earnings', 'number'), ('abs', 'Absence deduction', 'number'), ('absd', 'Absent days', 'number'),
          ('od', 'Other deductions', 'number'), ('rem', 'Remark', 'text')]
@app.route('/main')
@need()
def main():
    P = cur()
    if not P: return page('Main Payroll', '<p>Create a payroll period first.</p>')
    s = S(); ed = editable(P); emps = q("select * from main_employees where status='ACTIVE' order by full_name"); rows, skipped = [], []
    for e in emps:
        if not is_due(e['payment_frequency'], P['week_number']): skipped.append(e['full_name']); continue
        i = inp(P['id'], e['id'])
        cells = [Markup(f'<input name={e["id"]}_{k} type={t} step=any value="{E(i[k] or "")}" {"" if ed else "disabled"} {"placeholder=" + str(s["sun_hours"]) if k == "sun" else ""}>') for k, _, t in FIELDS]
        rows.append([e['full_name'], e['payment_frequency'], *cells])
    inputs = f'<form method=post action="/main/save">{hp(P)}{T(["Employee", "Freq"] + [l for _, l, _ in FIELDS], rows)}{"<button class=p>Save inputs</button>" if ed else ""}</form>'
    due_e = [e for e in emps if is_due(e['payment_frequency'], P['week_number'])]
    TY = {'NORMAL': 'Normal after-hours', 'SUNDAY': 'Sunday / weekend', 'NIGHT': 'Night work'}; mu = {'NORMAL': s['ot_mult'], 'SUNDAY': s['sun_mult'], 'NIGHT': s['night_mult']}
    ro = [(x['name'], x['date'], TY[x['type']], x['ne'], x['ae'], x['hours'], n2(x['hours'] * gross_at(x['eid'], P['start_date']) / s['hours'] * mu[x['type']]),
           Markup(f'<form method=post action="/overtime/del/{x["tb"]}/{x["id"]}">{hp(P)}<button>×</button></form>') if ed else '') for x in ot_list(P)]
    if ed and due_e:
        otform = ('<form method=post action="/overtime/add" class=card>' + hp(P) + '<h3>Add overtime / Sunday / night work</h3><div class=g>' + fl('Employee', sel('emp', due_e)) + fl('Date', ip('date', 'date', value=P['start_date'], min=P['start_date'], max=P['end_date'], required='required'))
                  + fl('Type', '<select name=type><option value=NORMAL>Normal after-hours</option><option value=SUNDAY>Sunday / weekend</option><option value=NIGHT>Night work</option></select>')
                  + fl('Finished at (normal only)', ip('end', 'time'), 'hours counted after ' + s['normal_end']) + fl('Or type hours', ip('hours', 'number', step='any', min='0', style='width:110px'), 'Sunday: blank = ' + n2(s['sun_hours']) + ' h')
                  + '<button class=p>Add overtime</button></div><small>Add one entry per day. The system adds them up and applies the multipliers from Settings.</small></form>')
    else: otform = '<div class=empty>Overtime can be added once an employee is due for payment in this period and the period is open.</div>'
    otcard = otform + '<h3>Overtime entries this period</h3>' + T(['Employee', 'Date', 'Type', 'Normal end', 'Finished at', 'Hours', 'Pay', ''], ro)
    note = f'<p><small>Not paid in this period (by payment frequency): {E(", ".join(skipped)) or "none"}.</small></p>'
    mr = [r for r in main_rows(P) if r['due']]
    res = T(['Employee', 'Gross', 'Freq', 'Basic', 'Transport', 'OT h', 'OT pay', 'Sun h', 'Sun pay', 'Night h', 'Night pay', 'Commission', 'Other earn.', 'Absence', 'Other ded.', 'Net pay'],
            [(r['name'], n2(r['monthly_gross']), r['frequency'], n2(r['basic']), n2(r['transport']), r['ot_hours'], n2(r['ot_pay']), r['sun_hours'], n2(r['sun_pay']), r['night_hours'], n2(r['night_pay']), n2(r['commission']), n2(r['other_earnings']), n2(r['absence']), n2(r['other_deductions']), n2(r['net_pay'])) for r in mr],
            ('TOTAL', '', '', *[n2(sum(r[k] for r in mr)) for k in ('basic', 'transport')], '', n2(sum(r['ot_pay'] for r in mr)), '', n2(sum(r['sun_pay'] for r in mr)), '', n2(sum(r['night_pay'] for r in mr)),
             *[n2(sum(r[k] for r in mr)) for k in ('commission', 'other_earnings', 'absence', 'other_deductions', 'net_pay')]))
    return page(f'Main Payroll · {P["label"]}', f'{note}{otcard}<h3>Other pay inputs</h3>{inputs}<h3>Calculated payroll</h3>{res}<p><a href="/export/main?p={P["id"]}">⬇ Excel report</a></p>')
@app.route('/main/save', methods=['POST'])
@need('ACCOUNTANT')
def main_save():
    P = cur()
    if blocked(P): return back()
    s = S()
    for e in q("select * from main_employees where status='ACTIVE'"):
        if not is_due(e['payment_frequency'], P['week_number']): continue
        if not any(x.startswith(f'{e["id"]}_') for x in request.form): continue
        old = calc(e, P, s)['net_pay']; k = lambda n: request.form.get(f'{e["id"]}_{n}', '').strip(); a = (P['id'], e['id'])
        for t in RAW: db().execute(f'delete from {t} where period_id=? and employee_id=?', a)
        if fv(k('tr')) or fv(k('ex')) or fv(k('al')): db().execute('insert into transport_records(period_id,employee_id,transport,expense,other_allowance) values(?,?,?,?,?)', (*a, fv(k('tr')), fv(k('ex')), fv(k('al'))))
        if fv(k('com')) or k('cd') or k('co') or k('rem'): db().execute('insert into commission_records(period_id,employee_id,amount,description,customer_order,remark) values(?,?,?,?,?,?)', (*a, fv(k('com')), k('cd'), k('co'), k('rem')))
        if fv(k('abs')) or fv(k('absd')): db().execute('insert into absence_records(period_id,employee_id,absent_days,deduction_amount) values(?,?,?,?)', (*a, fv(k('absd')), fv(k('abs'))))
        if fv(k('oe')): db().execute('insert into other_earnings(period_id,employee_id,amount,description) values(?,?,?,?)', (*a, fv(k('oe')), ''))
        if fv(k('od')): db().execute('insert into other_deductions(period_id,employee_id,amount,description) values(?,?,?,?)', (*a, fv(k('od')), ''))
        db().commit(); new = calc(e, P, s)['net_pay']
        if new != old: audit('Main payroll net pay', f'{e["full_name"]} / {P["label"]}', old, new)
    touch(P); return back()

# ---------- settings / users / audit ----------
@app.route('/settings', methods=['GET', 'POST'])
@need()
def settings():
    if request.method == 'POST':
        if me()['role'] != 'ADMIN': return page('Forbidden', ''), 403
        cur_ = S()
        for k in DEFAULTS:
            v = request.form.get(k, '').strip()
            if v and str(v) != str(cur_[k]) and (k == 'normal_end' or fv(v) > 0):
                ex('update settings set value=? where key=?', (v, k)); audit('Setting', k, cur_[k], v, request.form.get('reason', ''))
        return back()
    s = S(); dis = '' if me()['role'] == 'ADMIN' else 'disabled'
    f = ''.join(f'<p><label>{LABELS[k]}<br><input name={k} value="{s[k]}" {dis}></label></p>' for k in DEFAULTS)
    return page('Payroll Settings', f'<form method=post class=card>{f}<input name=reason placeholder="Reason (optional)" {dis}> <button class=p {dis}>Save settings</button></form><p><small>Only an administrator can change these. Already-calculated periods keep their saved figures.</small></p>' + ('<p><a href="/backup">⬇ Download database backup</a></p>' if me()['role'] == 'ADMIN' else ''))
@app.route('/users', methods=['GET', 'POST'])
@need('ADMIN')
def users():
    if request.method == 'POST':
        f = request.form
        if f['role'] not in ROLES or len(f['password']) < 6: flash('Role required and password must be 6+ characters.'); return back()
        try: ex('insert into users(username,full_name,password_hash,role) values(?,?,?,?)', (f['username'].strip(), f['name'], generate_password_hash(f['password']), f['role'])); audit('User added', f['username'], '', f['role'])
        except sqlite3.IntegrityError: flash('Username exists.')
        return back()
    form = '<form method=post class=card><h3>Add user</h3><div class=g>' + fl('Username', ip('username', required='required')) + fl('Full name', ip('name')) + fl('Password (6+ characters)', ip('password', 'password', required='required')) + fl('Role', '<select name=role>' + ''.join('<option>' + r + '</option>' for r in ROLES) + '</select>') + '<button class=p>Add user</button></div></form>'
    ro = lambda cur_: ''.join(f'<option {"selected" if r == cur_ else ""}>{r}</option>' for r in ROLES)
    rows = [(u['username'], u['full_name'], Markup(f'<form method=post action="/users/{u["id"]}/update" class=g><select name=role>{ro(u["role"])}</select><input name=password type=password placeholder="New password (optional)"><button>Save</button></form>'),
             'Active' if u['active'] else 'Inactive', Markup(f'<form method=post action="/users/{u["id"]}/toggle"><button>Activate/Deactivate</button></form>')) for u in q('select * from users')]
    return page('Users & Roles', form + T(['Username', 'Name', 'Role / reset password', 'Status', ''], rows))
@app.route('/users/<int:i>/update', methods=['POST'])
@need('ADMIN')
def user_update(i):
    f = request.form; u = q1('select * from users where id=?', (i,))
    if f['role'] in ROLES and f['role'] != u['role']:
        if i == me()['id']: flash('You cannot change your own role.')
        else: ex('update users set role=? where id=?', (f['role'], i)); audit('User role', u['username'], u['role'], f['role'])
    if f['password']:
        if len(f['password']) < 6: flash('Password must be 6+ characters.')
        else: ex('update users set password_hash=? where id=?', (generate_password_hash(f['password']), i)); audit('Password reset', u['username'])
    return back()
@app.route('/users/<int:i>/toggle', methods=['POST'])
@need('ADMIN')
def user_toggle(i):
    if i == me()['id']: flash('You cannot deactivate yourself.'); return back()
    ex('update users set active=1-active where id=?', (i,)); audit('User toggled', i); return back()
@app.route('/account', methods=['GET', 'POST'])
@need()
def account():
    if request.method == 'POST':
        u = q1('select * from users where id=?', (me()['id'],)); f = request.form
        if not check_password_hash(u['password_hash'], f['old']) or len(f['new']) < 6 or f['new'] != f['again']: flash('Wrong current password, or new passwords do not match / under 6 characters.')
        else: ex('update users set password_hash=? where id=?', (generate_password_hash(f['new']), u['id'])); audit('Password changed', u['username']); flash('Password changed.')
        return back()
    return page('My password', '<form method=post class=card><div class=g>' + fl('Current password', ip('old', 'password')) + fl('New password', ip('new', 'password')) + fl('Repeat new password', ip('again', 'password')) + '<button class=p>Change password</button></div></form>')
@app.route('/backup')
@need('ADMIN')
def backup():
    audit('Backup download', 'database'); return send_file(DB, as_attachment=True, download_name=f'markid-backup-{dt.date.today().isoformat()}.db')
@app.route('/adjustments', methods=['GET', 'POST'])
@need()
def adjustments():
    if request.method == 'POST':
        if me()['role'] not in ('ACCOUNTANT', 'ADMIN'): return page('Forbidden', ''), 403
        f = request.form; kind, pid_ = f['who'].split(':'); amt = fv(f['amount']); Tp = q1('select * from payroll_periods where id=?', (f['target'],))
        if not f['reason'].strip() or not amt: flash('Amount and reason are required.'); return back()
        if not Tp or Tp['status'] not in ('DRAFT', 'CALCULATED'): flash('Target period must be open (DRAFT or CALCULATED).'); return back()
        who = q1('select full_name n from contract_workers where id=?' if kind == 'CONTRACT' else 'select full_name n from main_employees where id=?', (int(pid_),))['n']
        ex('insert into payroll_adjustments(period_id,kind,person_id,amount,description,ref_period_id,reason,created_by,created_at) values(?,?,?,?,?,?,?,?,?)',
           (Tp['id'], kind, int(pid_), amt, f['desc'], f.get('ref') or None, f['reason'], me()['username'], dt.datetime.now().strftime('%Y-%m-%d %H:%M')))
        touch(Tp); audit('Adjustment', f'{who} / {Tp["label"]}', '', amt, f['reason']); return back()
    ws = q('select * from contract_workers where status="ACTIVE" order by full_name'); es = q('select * from main_employees where status="ACTIVE" order by full_name')
    who = ''.join(f'<option value="CONTRACT:{w["id"]}">Contract · {E(w["full_name"])}</option>' for w in ws) + ''.join(f'<option value="MAIN:{e["id"]}">Main · {E(e["full_name"])}</option>' for e in es)
    op = ''.join(f'<option value={x["id"]}>{E(x["label"])}</option>' for x in periods() if x['status'] in ('DRAFT', 'CALCULATED'))
    allp = '<option value="">—</option>' + ''.join(f'<option value={x["id"]}>{E(x["label"])} ({x["status"]})</option>' for x in periods())
    form = ('<form method=post class=card><h3>Add adjustment / other earning</h3><div class=g>' + fl('Person', f'<select name=who>{who}</select>') + fl('Amount (+ earning, − deduction)', ip('amount', 'number', step='any', required='required')) + fl('Description', ip('desc')) + fl('Apply in period', f'<select name=target>{op}</select>') + fl('Correcting period (optional)', f'<select name=ref>{allp}</select>') + fl('Reason (required)', ip('reason', size='28')) + '<button class=p>Add</button></div>'
            '<small>Contract workers: added to the period payable as Other earnings. Main employees: added to Other earnings (or Other deductions if negative). Use this to correct a LOCKED period.</small></form>')
    names = {}
    for w in q('select * from contract_workers'): names[('CONTRACT', w['id'])] = w['full_name']
    for e in q('select * from main_employees'): names[('MAIN', e['id'])] = e['full_name']
    lab = {x['id']: x['label'] for x in periods()}
    rows = [(r['created_at'], r['kind'], names.get((r['kind'], r['person_id']), '?'), n2(r['amount']), r['description'], lab.get(r['period_id']), lab.get(r['ref_period_id'], ''), r['reason'], r['created_by']) for r in q('select * from payroll_adjustments order by id desc')]
    return page('Adjustments', form + T(['When', 'Type', 'Person', 'Amount', 'Description', 'Applied in', 'Corrects', 'Reason', 'By'], rows))
@app.route('/audit')
@need('ADMIN', 'MANAGER', 'CHECKER', 'APPROVER', 'ACCOUNTANT')
def audit_page():
    return page('Audit log', T(['When', 'Who', 'What', 'Entity', 'Old', 'New', 'Reason'], [(r['at'], r['user'], r['action'], r['entity'], r['old_value'], r['new_value'], r['reason']) for r in q('select * from payroll_audit_logs order by id desc limit 500')]))

# ---------- overtime entries ----------
def ot_list(P):
    out = []
    for r in q('select o.*, e.full_name n from overtime_records o join main_employees e on e.id=o.employee_id where o.period_id=?', (P['id'],)):
        out.append(dict(tb='o', id=r['id'], eid=r['employee_id'], name=r['n'], date=r['work_date'] or '', type=r['ot_type'], ne=r['normal_end'] or '', ae=r['actual_end'] or '', hours=r['hours']))
    for r in q('select o.*, e.full_name n from night_work_records o join main_employees e on e.id=o.employee_id where o.period_id=?', (P['id'],)):
        out.append(dict(tb='n', id=r['id'], eid=r['employee_id'], name=r['n'], date=r['work_date'] or '', type='NIGHT', ne='', ae='', hours=r['hours']))
    return sorted(out, key=lambda x: (x['date'], x['name']))
@app.route('/overtime/add', methods=['POST'])
@need('ACCOUNTANT')
def ot_add():
    P = cur()
    if blocked(P): return back()
    f = request.form; s = S(); e = q1('select * from main_employees where id=?', (f.get('emp', 0),)); typ = f.get('type', 'NORMAL'); d = f.get('date', '').strip(); end = f.get('end', '').strip()
    if not e or not is_due(e['payment_frequency'], P['week_number']): flash('This employee is not paid in this period (see payment frequency and week number).'); return back()
    if not d or not (P['start_date'] <= d <= P['end_date']): flash(f'Date must be inside the period ({P["start_date"]} to {P["end_date"]}).'); return back()
    h = (minutes(end) - minutes(s['normal_end'])) / 60 if (typ == 'NORMAL' and end) else fv(f.get('hours'))
    if typ == 'SUNDAY' and not h: h = s['sun_hours']
    if h <= 0: flash('Overtime hours are zero. For normal overtime the finish time must be later than ' + s['normal_end'] + ', or type the hours.'); return back()
    h = round(h, 2)
    if typ == 'NIGHT': ex('insert into night_work_records(period_id,employee_id,hours,work_date) values(?,?,?,?)', (P['id'], e['id'], h, d))
    else: ex('insert into overtime_records(period_id,employee_id,ot_type,normal_end,actual_end,hours,work_date) values(?,?,?,?,?,?,?)', (P['id'], e['id'], typ if typ in ('NORMAL', 'SUNDAY') else 'NORMAL', s['normal_end'] if typ == 'NORMAL' else '', end if typ == 'NORMAL' else '', h, d))
    audit('Overtime added', f'{e["full_name"]} / {P["label"]}', '', f'{typ} {h} h on {d}'); touch(P); return back()
@app.route('/overtime/del/<tb>/<int:i>', methods=['POST'])
@need('ACCOUNTANT')
def ot_del(tb, i):
    P = cur()
    if tb not in ('o', 'n') or blocked(P): return back()
    t = 'overtime_records' if tb == 'o' else 'night_work_records'; r = q1(f'select * from {t} where id=? and period_id=?', (i, P['id']))
    if r:
        nm_ = q1('select full_name n from main_employees where id=?', (r['employee_id'],))['n']
        audit('Overtime deleted', f'{nm_} / {P["label"]}', f'{r["hours"]} h', ''); ex(f'delete from {t} where id=?', (i,)); touch(P)
    return back()

# ---------- excel ----------
HF = PatternFill('solid', fgColor='1F3A5F'); BF = PatternFill('solid', fgColor='DCE6F5'); TF = PatternFill('solid', fgColor='FFF2CC')
_t = Side(style='thin', color='BFBFBF'); BD = Border(left=_t, right=_t, top=_t, bottom=_t); MONEY = '#,##0.00'
def put(ws, r, c, v, bold=False, fill=None, fmt=None, color=None, align=None, border=True):
    x = ws.cell(r, c, v)
    if bold or color: x.font = Font(bold=bold, color=color)
    if fill: x.fill = fill
    if fmt: x.number_format = fmt
    if border: x.border = BD
    if align: x.alignment = Alignment(horizontal=align, vertical='center', wrap_text=True)
    return x
def hdr(ws, r, labels, c0=1):
    for j, l in enumerate(labels): put(ws, r, c0 + j, l, True, HF, color='FFFFFF', align='center')
def head(ws, text, sub):
    ws.cell(1, 1, text).font = Font(bold=True, size=14); ws.cell(2, 1, sub).font = Font(italic=True, color='555555')
def widths(ws, ws_):
    for i, w in enumerate(ws_, 1): ws.column_dimensions[chr(64 + i)].width = w
def sumrow(ws, r, c_from, c_to, first, last, skip=(), label='TOTAL'):
    for j in range(1, c_to + 1): put(ws, r, j, '', fill=TF)
    put(ws, r, 2, label, True, TF)
    for j in range(c_from, c_to + 1):
        if j not in skip: put(ws, r, j, f'=SUM({chr(64 + j)}{first}:{chr(64 + j)}{last})' if last >= first else 0, True, TF, MONEY)
def sign_block(ws, r):
    for i, l in enumerate(('Prepared By', 'Checked By', 'Approved By')):
        put(ws, r + i * 2, 2, f'{l}: ____________________', border=False); put(ws, r + i * 2, 4, 'Sign: ______________', border=False); put(ws, r + i * 2, 6, 'Date: ____________', border=False)
def send(wb, name):
    b = io.BytesIO(); wb.save(b); b.seek(0); return send_file(b, as_attachment=True, download_name=name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
WORK_SQL = 'select dd.name dn, dd.category cat, i.quantity, i.unit_price, c.name cn, i.remark from contract_work_items i join contract_work cw on cw.id=i.work_id join designs dd on dd.id=i.design_id left join customers c on c.id=i.customer_id where cw.worker_id=? and cw.period_id=? order by i.id'
@app.route('/export/main')
@need()
def export_main():
    P = cur(); s = S(); rs = [r for r in main_rows(P) if r['due']]; wb = Workbook(); ws = wb.active; ws.title = 'Payroll Report'; sub = f'{P["label"]} · {P["start_date"]} to {P["end_date"]} · Status: {P["status"]}'
    head(ws, 'Markid Furniture — Main Employees Payroll', sub); widths(ws, [6, 28, 14, 14, 16, 14, 14, 12, 14, 14, 14, 30, 16])
    hdr(ws, 4, ['NO', 'Name of employees', 'Gross salary', 'Basic salary', 'Transport/Expense', 'Overtime', 'Night work', 'Absent', 'Commission', 'Other earnings', 'Other deductions', 'Remark', 'Net pay']); r = 4
    for n, x in enumerate(rs, 1):
        r += 1
        for j, v in enumerate([n, x['name'], x['monthly_gross'], x['basic'], x['transport'], round(x['ot_pay'] + x['sun_pay'], 2), x['night_pay'], x['absence'], x['commission'], x['other_earnings'], x['other_deductions'], x['remark'], x['net_pay']], 1): put(ws, r, j, v, fmt=None if j in (1, 2, 12) else MONEY)
    r += 1; sumrow(ws, r, 3, 13, 5, r - 1, skip=(12,)); sign_block(ws, r + 3)
    b = wb.create_sheet('Pay Breakdown'); head(b, 'Pay breakdown per employee', sub); widths(b, [5, 26, 12, 14, 12, 14, 11, 14, 10, 14, 10, 14, 15, 14, 30, 14, 11, 14, 14, 15, 28])
    hdr(b, 4, ['NO', 'Employee', 'Frequency', 'Monthly gross', 'Hourly rate', 'Basic (period)', 'Normal OT hrs', 'Normal OT pay', 'Sunday hrs', 'Sunday pay', 'Night hrs', 'Night pay', 'Transport / expense', 'Commission', 'Commission description / order', 'Other earnings', 'Absent days', 'Absence deduction', 'Other deductions', 'Net pay', 'Remark']); r = 4
    for n, x in enumerate(rs, 1):
        i = inp(P['id'], x['employee_id']); r += 1
        for j, v in enumerate([n, x['name'], x['frequency'], x['monthly_gross'], x['monthly_gross'] / s['hours'], x['basic'], x['ot_hours'], x['ot_pay'], x['sun_hours'], x['sun_pay'], x['night_hours'], x['night_pay'], x['transport'], x['commission'], ' / '.join(t for t in (i['cd'], i['co']) if t),
                               x['other_earnings'], i['absd'], x['absence'], x['other_deductions'], x['net_pay'], x['remark']], 1): put(b, r, j, v, fmt=None if j in (1, 2, 3, 15, 21) else MONEY)
    r += 1; sumrow(b, r, 4, 20, 5, r - 1, skip=(5, 15)); b.freeze_panes = 'C5'
    o = wb.create_sheet('Overtime Entries'); head(o, 'Overtime / Sunday / night entries (day by day)', sub); widths(o, [5, 26, 13, 20, 12, 12, 10, 13, 11, 14])
    hdr(o, 4, ['NO', 'Employee', 'Date', 'Type', 'Normal end', 'Finished at', 'Hours', 'Hourly rate', 'Multiplier', 'Pay']); r = 4; mu = {'NORMAL': s['ot_mult'], 'SUNDAY': s['sun_mult'], 'NIGHT': s['night_mult']}; TY = {'NORMAL': 'Normal after-hours', 'SUNDAY': 'Sunday / weekend', 'NIGHT': 'Night work'}
    for n, x in enumerate(ot_list(P), 1):
        rate = gross_at(x['eid'], P['start_date']) / s['hours']; r += 1
        for j, v in enumerate([n, x['name'], x['date'], TY[x['type']], x['ne'], x['ae'], x['hours'], rate, mu[x['type']], x['hours'] * rate * mu[x['type']]], 1): put(o, r, j, v, fmt=MONEY if j in (8, 10) else None)
    r += 1; sumrow(o, r, 7, 10, 5, r - 1, skip=(8, 9))
    audit('Excel export', 'main', '', P['label']); return send(wb, f'markid_main_payroll_{P["id"]}.xlsx')
@app.route('/export/contract')
@need()
def export_contract():
    P = cur(); cr = contract_rows(P); wb = Workbook(); ws = wb.active; ws.title = 'Payroll Report'; sub = f'{P["label"]} · {P["start_date"]} to {P["end_date"]} · Status: {P["status"]}'
    head(ws, 'Markid Furniture — Contract Workers Payroll', sub); widths(ws, [6, 30, 18, 16, 16, 16, 20, 22, 36])
    hdr(ws, 4, ['NO', 'Worker Name', 'Previous Balance', 'Current Work', 'Other Earnings', 'Total Payable', 'Amount Paid', 'Remaining Balance', 'Remark']); r = 4
    for n, (w, b) in enumerate(cr, 1):
        r += 1
        for j, v in enumerate([n, w['full_name'], b['prev'], b['earned'], b['other'], b['payable'], b['paid'], b['close'], w['notes'] or ''], 1): put(ws, r, j, v, fmt=MONEY if 3 <= j <= 8 else None)
    r += 1; sumrow(ws, r, 3, 8, 5, r - 1); put(ws, r, 9, '', fill=TF)
    r += 2; ws.cell(r, 1, 'DETAILED WORK DONE — BY WORKER').font = Font(bold=True, size=12); grand = 0
    for n, (w, b) in enumerate(cr, 1):
        r += 2
        for j in range(1, 10): put(ws, r, j, '', fill=BF)
        put(ws, r, 1, n, True, BF); put(ws, r, 2, w['full_name'], True, BF); put(ws, r, 3, 'Phone: ' + (w['phone'] or '-'), fill=BF); put(ws, r, 8, 'Notes: ' + (w['notes'] or '-'), fill=BF)
        r += 1; hdr(ws, r, ['Design', 'Category', 'Quantity', 'Worker Price', 'Total', 'Customer', 'Remark'], 2); ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        items = q(WORK_SQL, (w['id'], P['id']))
        for it in items:
            r += 1
            for j, v in enumerate([it['dn'], it['cat'] or '', it['quantity'], it['unit_price'], it['quantity'] * it['unit_price'], it['cn'] or '', it['remark'] or ''], 2): put(ws, r, j, v, fmt=MONEY if j in (5, 6) else None)
            put(ws, r, 9, ''); ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        if not items: r += 1; put(ws, r, 2, 'No work entered this period', border=False)
        for lab, v, hi in (('Previous unpaid balance', b['prev'], False), ('Current work total', b['earned'], False), ('Other earnings / adjustments', b['other'], False), ('Total payable', b['payable'], True), ('Amount paid', b['paid'], False), ('Remaining balance', b['close'], True)):
            r += 1; f_ = TF if hi else None; put(ws, r, 4, lab, True, f_, align='right'); put(ws, r, 5, '', fill=f_); ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5); put(ws, r, 6, v, True, f_, MONEY)
        grand += b['earned']
    r += 2; put(ws, r, 4, 'TOTAL CURRENT WORK (all workers)', True, TF, align='right'); put(ws, r, 5, '', fill=TF); ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5); put(ws, r, 6, grand, True, TF, MONEY); sign_block(ws, r + 3)
    a = wb.create_sheet('All Work Lines'); head(a, 'All work lines (filter by worker, design or customer)', sub); widths(a, [26, 30, 22, 16, 11, 14, 15, 22, 36]); hdr(a, 4, ['Period', 'Worker', 'Design', 'Category', 'Quantity', 'Worker Price', 'Total', 'Customer', 'Remark']); r = 4
    for w, b in cr:
        for it in q(WORK_SQL, (w['id'], P['id'])):
            r += 1
            for j, v in enumerate([P['label'], w['full_name'], it['dn'], it['cat'] or '', it['quantity'], it['unit_price'], it['quantity'] * it['unit_price'], it['cn'] or '', it['remark'] or ''], 1): put(a, r, j, v, fmt=MONEY if j in (6, 7) else None)
    r += 1; sumrow(a, r, 5, 7, 5, r - 1, skip=(6,)); a.auto_filter.ref = f'A4:I{max(r - 1, 5)}'; a.freeze_panes = 'A5'
    d = wb.create_sheet('By Design'); head(d, 'Production by design', sub); widths(d, [30, 20, 16, 18, 14]); hdr(d, 4, ['Design', 'Category', 'Total Quantity', 'Total Value', 'Workers']); r = 4
    for x in q('select dd.name n, dd.category c, sum(i.quantity) qy, sum(i.quantity*i.unit_price) v, count(distinct cw.worker_id) wk from contract_work_items i join contract_work cw on cw.id=i.work_id join designs dd on dd.id=i.design_id where cw.period_id=? group by dd.id order by v desc', (P['id'],)):
        r += 1
        for j, v in enumerate([x['n'], x['c'] or '', x['qy'], x['v'], x['wk']], 1): put(d, r, j, v, fmt=MONEY if j == 4 else None)
    audit('Excel export', 'contract', '', P['label']); return send(wb, f'markid_contract_payroll_{P["id"]}.xlsx')
@app.route('/export/statement/<int:i>')
@need()
def export_statement(i):
    w = q1('select * from contract_workers where id=?', (i,)); wb = Workbook(); ws = wb.active; ws.title = 'Statement'
    head(ws, f'Account statement — {w["full_name"]}', f'Phone: {w["phone"] or "-"} · Start: {w["start_date"] or "-"}'); widths(ws, [6, 34, 16, 16, 16, 16, 16, 16]); hdr(ws, 4, ['NO', 'Period', 'Opening balance', 'Work earned', 'Other earnings', 'Total payable', 'Paid', 'Closing balance']); r = 4; n = 0
    for X in periods():
        b = bal(i, X['id'])
        if b['earned'] or b['other'] or b['paid'] or b['prev']:
            r += 1; n += 1
            for j, v in enumerate([n, X['label'], b['prev'], b['earned'], b['other'], b['payable'], b['paid'], b['close']], 1): put(ws, r, j, v, fmt=MONEY if j > 2 else None)
    sign_block(ws, r + 3)
    h = wb.create_sheet('Work History'); head(h, f'All work done — {w["full_name"]}', ''); widths(h, [34, 30, 22, 11, 14, 15, 22, 36]); hdr(h, 4, ['Period', 'Design', 'Category', 'Quantity', 'Worker Price', 'Total', 'Customer', 'Remark']); r = 4
    for X in periods():
        for it in q(WORK_SQL, (i, X['id'])):
            r += 1
            for j, v in enumerate([X['label'], it['dn'], it['cat'] or '', it['quantity'], it['unit_price'], it['quantity'] * it['unit_price'], it['cn'] or '', it['remark'] or ''], 1): put(h, r, j, v, fmt=MONEY if j in (5, 6) else None)
    p = wb.create_sheet('Payments'); head(p, f'Payments — {w["full_name"]}', ''); widths(p, [34, 16, 16, 20]); hdr(p, 4, ['Period', 'Amount paid', 'Paid by', 'Paid at']); r = 4
    for x in q('select p.*, pp.label from contract_worker_payments p join payroll_periods pp on pp.id=p.period_id where worker_id=? order by pp.start_date', (i,)):
        r += 1
        for j, v in enumerate([x['label'], x['amount'], x['paid_by'], x['paid_at']], 1): put(p, r, j, v, fmt=MONEY if j == 2 else None)
    return send(wb, f'statement_{w["full_name"].replace(" ", "_")}.xlsx')

# ---------- offline sync (server <-> desktop clients) ----------
# Server = source of truth. A desktop client records every successful write as an operation (with ids replaced by stable uids),
# keeps working offline, then PUSHES the operations (server replays them with the same rules/permissions) and PULLS a fresh snapshot.
TOKEN_RX = re.compile(r'@([pwdeion]):([0-9a-f]{32})@')
SKIP_RX = re.compile(r'^/(login|logout|account|users|api|backup|sync)')
IDF = {'p': 'p', 'target': 'p', 'ref': 'p', 'w': 'w', 'd': 'd', 'design': 'd', 'emp': 'e'}
PATHRX = [(re.compile(r'^(/period/)(\d+)'), 'p'), (re.compile(r'^(/designs/)(\d+)'), 'd'), (re.compile(r'^(/workers/)(\d+)'), 'w'), (re.compile(r'^(/contract/del/)(\d+)'), 'i'), (re.compile(r'^(/employees/)(\d+)'), 'e')]
SUMMARIES = [(r'^/contract/add', 'Added contract work'), (r'^/contract/pay', 'Saved contract payments'), (r'^/contract/del', 'Deleted a contract work entry'), (r'^/overtime/add', 'Added overtime'),
             (r'^/overtime/del', 'Deleted an overtime entry'), (r'^/main/save', 'Saved main payroll inputs'), (r'^/workers/.*/price', 'Set a worker design price'), (r'^/workers/.+', 'Edited a worker'),
             (r'^/workers$', 'Added a worker'), (r'^/designs/.*/toggle', 'Activated/deactivated a design'), (r'^/designs$', 'Added a design'), (r'^/employees/.*/salary', 'Changed an employee salary'),
             (r'^/employees/.+', 'Edited an employee'), (r'^/employees$', 'Added an employee'), (r'^/adjustments', 'Added an adjustment'), (r'^/periods', 'Created a payroll period'),
             (r'^/period/.*/advance', 'Moved a period forward'), (r'^/period/.*/reject', 'Sent a period back'), (r'^/period/.*/unlock', 'Unlocked a period'), (r'^/settings', 'Changed settings')]
def summarize(path):
    for rx, t in SUMMARIES:
        if re.match(rx, path): return t
    return path
def maxrows(): return {t: q1(f'select coalesce(max(id),0) m from {t}')['m'] for t in UT.values()}
def tok(code, i):
    r = q1(f'select uid from {UT[code]} where id=?', (int(i),))
    return f'@{code}:{r["uid"]}@' if r and r['uid'] else str(i)
def tokenize(path, form):
    for rx, code in PATHRX:
        m = rx.match(path)
        if m: path = m.group(1) + tok(code, m.group(2)) + path[m.end():]; break
    m = re.match(r'^(/overtime/del/([on])/)(\d+)', path)
    if m: path = m.group(1) + tok(m.group(2), m.group(3)) + path[m.end():]
    out = {}
    for k, v in form.items():
        if k in IDF and v.isdigit(): v = tok(IDF[k], v)
        elif k == 'who' and ':' in v:
            kind, i = v.split(':', 1)
            if i.isdigit(): v = kind + ':' + tok('w' if kind == 'CONTRACT' else 'e', i)
        m = re.match(r'^pay_(\d+)$', k)
        if m: k = 'pay_' + tok('w', m.group(1))
        m = re.match(r'^(\d+)_(\w+)$', k)
        if m: k = tok('e', m.group(1)) + '_' + m.group(2)
        out[k] = v
    return path, out
def sdb():
    c = sqlite3.connect(STATE_DB); c.row_factory = sqlite3.Row
    c.executescript('create table if not exists outbox(seq integer primary key, op_id text, user text, at text, path text, form text, created text, summary text);'
                    'create table if not exists conflicts(id integer primary key, at text, summary text, error text, user text);create table if not exists config(key text primary key, value text);')
    return c
def cfg(k=None, v=None):
    c = sdb()
    if v is not None: c.execute('insert or replace into config values(?,?)', (k, v)); c.commit()
    d = {r['key']: r['value'] for r in c.execute('select * from config')}; c.close(); return d.get(k, '') if k else d
@app.before_request
def _pre():
    if MODE == 'client' and request.method == 'POST' and re.match(r'^/(users|account)', request.path):
        flash('Users and passwords are managed on the online server. Do this there, then sync.'); return redirect('/sync')
    if MODE == 'client' and request.method == 'POST' and not SKIP_RX.match(request.path) and me():
        if SYNCING: return 'Synchronising with the server - try again in a few seconds.', 503
        g._pre = (maxrows(), len(session.get('_flashes', [])))
@app.after_request
def _rec(resp):
    try:
        if hasattr(g, '_pre') and resp.status_code < 400 and me():
            pre, nfl = g._pre
            if len(session.get('_flashes', [])) <= nfl: _record_op(pre)
    except Exception: app.logger.exception('could not record sync operation')
    return resp
def _record_op(pre):
    form = {k: v for k, v in request.form.items() if k != '_csrf'}
    if request.path.startswith(('/contract', '/main', '/overtime')) and 'p' not in form:
        P = cur()
        if P: form['p'] = str(P['id'])
    path, form = tokenize(request.path, form)
    created = [[code, r['uid']] for code, t in UT.items() for r in q(f'select uid from {t} where id > ? order by id', (pre[t],))]
    c = sdb(); c.execute('insert into outbox(op_id,user,at,path,form,created,summary) values(?,?,?,?,?,?,?)',
                         (uuid.uuid4().hex, me()['username'], dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), path, json.dumps(form), json.dumps(created), summarize(path))); c.commit(); c.close()

# ----- server side -----
def _auth(): return MODE == 'server' and bool(SYNC_TOKEN) and hmac.compare_digest(request.headers.get('X-Sync-Token', ''), SYNC_TOKEN)
def replay(op):
    u = q1('select * from users where username=? and active=1', (op['user'],))
    if not u: return 'rejected', f'User "{op["user"]}" does not exist or is inactive on the server.'
    miss = []
    def sub(m):
        r = q1(f'select id from {UT[m.group(1)]} where uid=?', (m.group(2),))
        if not r: miss.append(1); return '0'
        return str(r['id'])
    path = TOKEN_RX.sub(sub, op['path']); form = {TOKEN_RX.sub(sub, k): TOKEN_RX.sub(sub, str(v)) for k, v in op['form'].items()}
    if miss: return 'rejected', 'Refers to a record that does not exist on the server (it was rejected earlier or removed).'
    pre = maxrows(); c = app.test_client()
    with c.session_transaction() as s: s['u'] = {'id': u['id'], 'username': u['username'], 'role': u['role']}; s['csrf'] = 'sync'
    form['_csrf'] = 'sync'
    try: r = c.post(path, data=form)
    except Exception as e: return 'rejected', f'Server error: {e}'
    with c.session_transaction() as s: fl = s.get('_flashes') or []
    if fl or r.status_code >= 400: return 'rejected', ' '.join(m for _, m in fl) or f'Refused (HTTP {r.status_code}).'
    for t in UT.values():
        exp = [uid for code, uid in op.get('created', []) if UT.get(code) == t]
        for row, uid in zip(q(f'select id from {t} where id > ? order by id', (pre[t],)), exp):
            try: ex(f'update {t} set uid=? where id=?', (uid, row['id']))
            except sqlite3.IntegrityError: pass
    return 'applied', ''
@app.route('/api/sync/ping')
def api_ping():
    if not _auth(): return {'error': 'forbidden'}, 403
    return {'ok': True, 'schema': SCHEMA_VERSION, 'version': VERSION}
@app.route('/api/sync/push', methods=['POST'])
def api_push():
    if not _auth(): return {'error': 'forbidden'}, 403
    d = request.get_json(force=True); dev = str(d.get('device') or 'unknown')[:60]
    if d.get('schema') != SCHEMA_VERSION: return {'error': 'The server runs a different app version. Update the app.'}, 409
    results = []
    for op in d.get('ops', []):
        if q1('select 1 from sync_applied where op_id=?', (op['op_id'],)): results.append({'op_id': op['op_id'], 'status': 'duplicate', 'message': ''}); continue
        st, msg = replay(op)
        ex('insert into sync_applied(op_id,device,user,at,summary,status,message) values(?,?,?,?,?,?,?)', (op['op_id'], dev, op['user'], op.get('at', ''), op.get('summary', ''), st.upper(), msg))
        results.append({'op_id': op['op_id'], 'status': st, 'message': msg})
    rej = sum(1 for r in results if r['status'] == 'rejected')
    ex('insert into sync_devices(device,last_push,ops_pushed,ops_rejected) values(?,?,?,?) on conflict(device) do update set last_push=excluded.last_push, ops_pushed=ops_pushed+excluded.ops_pushed, ops_rejected=ops_rejected+excluded.ops_rejected',
       (dev, dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), len(results), rej))
    return {'results': results}
@app.route('/api/sync/pull')
def api_pull():
    if not _auth(): return {'error': 'forbidden'}, 403
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.db'); tmp.close()
    src = sqlite3.connect(DB); dst = sqlite3.connect(tmp.name); src.backup(dst); dst.close(); src.close()
    data = open(tmp.name, 'rb').read(); os.remove(tmp.name); dev = request.args.get('device', 'unknown')[:60]
    ex('insert into sync_devices(device,last_pull) values(?,?) on conflict(device) do update set last_pull=excluded.last_pull', (dev, dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    return send_file(io.BytesIO(data), mimetype='application/octet-stream', download_name='markid.db')

# ----- client side -----
def _http(url, token, data=None, timeout=90):
    req = urllib.request.Request(url, data=json.dumps(data).encode() if data is not None else None, headers={'X-Sync-Token': token, 'Content-Type': 'application/json'})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r: return r.read()
    except urllib.error.HTTPError as e:
        if e.code == 403: raise RuntimeError('wrong sync token')
        try: raise RuntimeError(json.loads(e.read()).get('error', f'server answered {e.code}'))
        except ValueError: raise RuntimeError(f'server answered {e.code}')
def replace_db(data):
    tmp = DB + '.incoming'; open(tmp, 'wb').write(data)
    d = g.pop('db', None)
    if d: d.close()
    for _ in range(20):
        try: os.replace(tmp, DB); break
        except PermissionError: time.sleep(0.5)
    else: raise RuntimeError('could not replace the local database (is it open elsewhere?)')
    for ext in ('-journal', '-wal', '-shm'):
        if os.path.exists(DB + ext): os.remove(DB + ext)
def sync_now():
    global SYNCING
    c0 = cfg(); url = (c0.get('url') or '').rstrip('/'); token = c0.get('token') or ''; dev = c0.get('device') or socket.gethostname()
    if not url or not token: return False, 'Enter the server address and sync token first.'
    if SYNCING: return False, 'A sync is already running.'
    SYNCING = True
    try:
        try: info = json.loads(_http(url + '/api/sync/ping', token, timeout=15))
        except Exception as e: return False, f'Cannot reach the server ({e}). You can keep working offline; it will sync later.'
        if info.get('schema') != SCHEMA_VERSION: return False, 'The server runs a different app version. Install the latest app.'
        c = sdb(); ops = c.execute('select * from outbox order by seq').fetchall(); pushed = rejected = 0
        if ops:
            payload = {'device': dev, 'schema': SCHEMA_VERSION, 'ops': [{'op_id': o['op_id'], 'user': o['user'], 'at': o['at'], 'path': o['path'], 'form': json.loads(o['form']), 'created': json.loads(o['created']), 'summary': o['summary']} for o in ops]}
            res = {r['op_id']: r for r in json.loads(_http(url + '/api/sync/push', token, payload, timeout=300))['results']}
            if any(o['op_id'] not in res for o in ops): return False, 'The server did not confirm every change. Nothing was lost; try again.'
            for o in ops:
                r = res[o['op_id']]; c.execute('delete from outbox where seq=?', (o['seq'],)); pushed += 1
                if r['status'] == 'rejected': rejected += 1; c.execute('insert into conflicts(at,summary,error,user) values(?,?,?,?)', (o['at'], o['summary'], r['message'], o['user']))
            c.commit()
        c.close()
        data = _http(url + '/api/sync/pull?device=' + urllib.parse.quote(dev), token, timeout=300)
        if data[:16] != b'SQLite format 3\x00': return False, 'The server sent an invalid database.'
        replace_db(data); stamp = dt.datetime.now().strftime('%Y-%m-%d %H:%M'); cfg('last_sync', 'Synced ' + stamp)
        return True, f'Synced at {stamp}: {pushed} change(s) sent' + (f', {rejected} could not be applied (see below)' if rejected else '') + '.'
    except Exception as e: return False, f'Sync failed: {e}'
    finally: SYNCING = False
def start_autosync(every=60):
    def loop():
        time.sleep(10)
        while True:
            try:
                with app.app_context():
                    if MODE == 'client' and cfg('auto') != '0' and cfg('url'): sync_now()
            except Exception: pass
            time.sleep(every)
    threading.Thread(target=loop, daemon=True).start()
@app.route('/sync')
@need()
def sync_page():
    u = me()
    if MODE == 'server':
        if u['role'] not in ('ADMIN', 'MANAGER'): return page('Forbidden', '<p class=err>Only an administrator or manager can see sync activity.</p>'), 403
        dv = T(['Device', 'Last sent changes', 'Last downloaded', 'Changes received', 'Rejected'], [(d['device'], d['last_push'] or '-', d['last_pull'] or '-', d['ops_pushed'], d['ops_rejected']) for d in q('select * from sync_devices order by device')])
        ops = T(['Device time', 'Device', 'User', 'What', 'Result', 'Message'], [(o['at'], o['device'], o['user'], o['summary'], o['status'], o['message']) for o in q('select * from sync_applied order by rowid desc limit 100')])
        ok = 'ready' if SYNC_TOKEN else 'NOT ready: set the SYNC_TOKEN environment variable'
        return page('Sync & devices', f'<div class=card><b>This is the online server.</b> Sync is {ok}. Desktop apps send their offline work here and download the latest data.</div><h3>Devices</h3>{dv}<h3>Recent changes received from devices</h3>{ops}')
    if MODE != 'client': return page('Sync & devices', '<div class=card>This copy runs <b>stand-alone</b> (no sync). Use the desktop app to work offline and sync with the online server.</div>')
    c = sdb(); pend = c.execute('select * from outbox order by seq').fetchall(); conf = c.execute('select * from conflicts order by id desc').fetchall(); c.close(); k = cfg()
    form = ('<form method=post action="/sync/config" class=card><h3>Connection to the online server</h3><div class=g>' + fl('Server address', ip('url', value=k.get('url', ''), size='34', placeholder='https://markid-payroll.onrender.com'))
            + fl('Sync token', ip('token', 'password', placeholder='saved' if k.get('token') else 'paste the token'), 'ask your administrator') + fl('This device name', ip('device', value=k.get('device', socket.gethostname())))
            + fl('Auto-sync every minute', f'<select name=auto><option value=1 {"selected" if k.get("auto") != "0" else ""}>On</option><option value=0 {"selected" if k.get("auto") == "0" else ""}>Off</option></select>') + '<button class=p>Save</button></div></form>')
    now = f'<form method=post action="/sync/now" class=card><div class=g><button class=p>🔄 Sync now</button><span><b>{len(pend)}</b> change(s) waiting · {E(k.get("last_sync") or "never synced")}</span></div><small>Working offline is fine: your entries are kept on this computer and sent automatically when the internet is back.</small></form>'
    pt = T(['Waiting since', 'User', 'What'], [(o['at'], o['user'], o['summary']) for o in pend])
    ct = T(['When', 'User', 'What', 'Why it was not applied', ''], [(o['at'], o['user'], o['summary'], o['error'], Markup(f'<form method=post action="/sync/dismiss"><input type=hidden name=id value={o["id"]}><button>Dismiss</button></form>')) for o in conf])
    return page('Sync & devices', now + form + '<h3>Waiting to be sent</h3>' + pt + ('<h3>Could not be applied (please enter again if still needed)</h3>' + ct if conf else '') + f'<p><small>Open in your normal browser instead: {E(request.host_url)}</small></p>')
@app.route('/sync/config', methods=['POST'])
@need('ACCOUNTANT')
def sync_config():
    f = request.form; cfg('url', f['url'].strip()); cfg('device', f['device'].strip() or socket.gethostname()); cfg('auto', f.get('auto', '1'))
    if f.get('token', '').strip(): cfg('token', f['token'].strip())
    flash('Saved.', 'ok'); return redirect('/sync')
@app.route('/sync/now', methods=['POST'])
@need('ACCOUNTANT')
def sync_now_route():
    ok, msg = sync_now(); flash(msg, 'ok' if ok else 'err'); return redirect('/sync')
@app.route('/sync/dismiss', methods=['POST'])
@need('ACCOUNTANT')
def sync_dismiss():
    c = sdb(); c.execute('delete from conflicts where id=?', (request.form.get('id'),)); c.commit(); c.close(); return redirect('/sync')


init_db()
if __name__ == '__main__':
    if MODE == 'client' and os.environ.get('WERKZEUG_RUN_MAIN') == 'true': start_autosync()
    app.run(host=os.environ.get('MARKID_HOST', '127.0.0.1'), port=int(os.environ.get('PORT', 5000)), use_reloader=True)
