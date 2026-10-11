"""End-to-end sync test: a real SERVER process + an offline CLIENT. Run: python test_sync.py"""
import os, sys, re, time, json, shutil, sqlite3, tempfile, subprocess, urllib.request, urllib.parse, http.cookiejar
T = tempfile.mkdtemp(); PORT = 5171; URL = f'http://127.0.0.1:{PORT}'
env = {**os.environ, 'MARKID_MODE': 'server', 'MARKID_DB': f'{T}/server.db', 'SYNC_TOKEN': 'tok123', 'ADMIN_PASSWORD': 'adminpw1', 'SECRET_KEY': 'k', 'PORT': str(PORT)}
srv = subprocess.Popen([sys.executable, '-c', f'import app; app.app.run(port={PORT})'], env=env, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
try:
    for _ in range(40):
        try: urllib.request.urlopen(URL + '/login'); break
        except Exception: time.sleep(0.5)
    # ---- web session on the SERVER (admin online) ----
    jar = http.cookiejar.CookieJar(); op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
    def sget(p): return op.open(URL + p).read().decode()
    def spost(path_, **d):
        t = re.search(r'name=_csrf value="(\w+)"', sget('/login')).group(1); d['_csrf'] = t
        return op.open(urllib.request.Request(URL + path_, urllib.parse.urlencode(d).encode())).read().decode()
    spost('/login', username='admin', password='adminpw1')
    spost('/users', username='acc', name='Acc', password='accpass1', role='ACCOUNTANT')
    spost('/periods', label='Wk1', week='1', start='2026-01-05', end='2026-01-11', cal='G.C.')
    spost('/designs', name='Addis Design', cat='', desc='', notes='')
    spost('/workers', name='Ambachew', phone='', start='', open='0', notes='')
    spost('/workers/1/price', design='1', price='900', eff='2026-01-01', notes='')
    spost('/employees', name='Emp One', dept='', pos='', gross='12000', freq='WEEKLY', start='', notes='')
    # ---- CLIENT (desktop app, in this process) ----
    os.environ.update(MARKID_MODE='client', MARKID_DB=f'{T}/client.db', MARKID_STATE=f'{T}/state.db', SECRET_KEY='c'); os.environ.pop('SYNC_TOKEN', None)
    import app as A
    c = A.app.test_client()
    def tok():
        c.get('/login')
        with c.session_transaction() as s: return s['csrf']
    def post(u, **d): d['_csrf'] = tok(); return c.post(u, data=d, follow_redirects=True)
    post('/login', username='admin', password='admin123')
    assert 'Sync' in c.get('/sync').get_data(as_text=True)
    # wrong token handled
    post('/sync/config', url=URL, token='WRONG', device='laptop-1', auto='1'); assert 'wrong sync token' in post('/sync/now').get_data(as_text=True)
    post('/sync/config', url=URL, token='tok123', device='laptop-1', auto='1')
    r = post('/sync/now'); assert 'Synced at' in r.get_data(as_text=True), re.findall(r'class="alert[^>]*>([^<]*)', r.get_data(as_text=True))
    post('/login', username='admin', password='adminpw1')                      # users now come from the server
    sd = lambda q, a=(): sqlite3.connect(f'{T}/server.db').execute(q, a).fetchall()
    cd = lambda q, a=(): sqlite3.connect(f'{T}/client.db').execute(q, a).fetchall()
    assert cd('select full_name from contract_workers') == [('Ambachew',)], 'pull failed'
    # ---- OFFLINE work: stop the server's reachability by working without syncing ----
    post('/contract/add', p='1', w='1', d='1', qty='6', cust='Sam', remark='offline entry')
    post('/workers', name='Offline Newbie', phone='', start='', open='0', notes='')             # new worker created offline
    nid = cd("select id from contract_workers where full_name='Offline Newbie'")[0][0]
    post(f'/workers/{nid}/price', design='1', price='700', eff='2026-01-01', notes='')           # refers to the offline-created worker
    post('/contract/add', p='1', w=str(nid), d='1', qty='2', cust='', remark='')
    post('/contract/pay', p='1', pay_1='1000', **{f'pay_{nid}': '0'})
    post('/overtime/add', p='1', emp='1', date='2026-01-06', type='NORMAL', end='20:30', hours='')
    post('/main/save', p='1', **{'1_tr': '2650'})
    n = sqlite3.connect(f'{T}/state.db').execute('select count(*) from outbox').fetchone()[0]; assert n >= 6, n
    assert sd('select count(*) from contract_work_items')[0][0] == 0                              # server has nothing yet
    assert 'waiting to sync' in c.get('/').get_data(as_text=True)
    # ---- SYNC ----
    r = post('/sync/now').get_data(as_text=True); assert 'Synced at' in r and '0 could not' not in r.replace('could not be applied', ''), r[-500:]
    assert sqlite3.connect(f'{T}/state.db').execute('select count(*) from outbox').fetchone()[0] == 0
    assert sd('select count(*) from contract_work_items')[0][0] == 2, sd('select * from contract_work_items')
    assert sd("select count(*) from contract_workers where full_name='Offline Newbie'")[0][0] == 1
    assert sd('select hours from overtime_records')[0][0] == 3.0 and sd('select transport from transport_records')[0][0] == 2650
    wuid_c = cd("select uid from contract_workers where full_name='Offline Newbie'")[0][0]; wuid_s = sd("select uid from contract_workers where full_name='Offline Newbie'")[0][0]; assert wuid_c == wuid_s
    assert sd('select price from contract_worker_design_prices where worker_id=2')[0][0] == 700.0   # price op referenced the offline-created worker
    assert cd('select count(*) from contract_work_items')[0][0] == 2
    # admin sees progress online
    page = sget('/sync'); assert 'laptop-1' in page and 'Added contract work' in page and 'APPLIED' in page
    # ---- idempotency: re-sending an already-applied op is ignored ----
    ops = sd('select op_id from sync_applied limit 1'); body = json.dumps({'device': 'laptop-1', 'schema': 1, 'ops': [{'op_id': ops[0][0], 'user': 'admin', 'path': '/contract/pay', 'form': {}, 'created': []}]}).encode()
    rq = urllib.request.Request(URL + '/api/sync/push', body, {'X-Sync-Token': 'tok123', 'Content-Type': 'application/json'}); assert 'duplicate' in urllib.request.urlopen(rq).read().decode()
    # ---- CONFLICT: server locks the period while the client is offline ----
    post('/contract/add', p='1', w='1', d='1', qty='1', cust='', remark='will conflict')
    for _ in range(5): spost('/period/1/advance', p='1')                                          # admin walks it to LOCKED
    r = post('/sync/now').get_data(as_text=True)
    assert 'could not be applied' in r and 'LOCKED' in r.upper() or 'editing is closed' in r, r[-700:]
    assert sd('select count(*) from contract_work_items')[0][0] == 2                              # the late entry was NOT applied
    assert cd('select count(*) from contract_work_items')[0][0] == 2                              # and the client now mirrors the server
    assert 'editing is closed' in sget('/sync') and 'REJECTED' in sget('/sync')
    print('SYNC TESTS OK')
finally:
    srv.terminate(); shutil.rmtree(T, ignore_errors=True)
