"""Run: python test_smoke.py  — checks the spec's worked examples on a temp DB."""
import os, tempfile
os.environ['MARKID_DB'] = os.path.join(tempfile.mkdtemp(), 't.db')
import app as A
class _Q:
    pass
_q1=A.q1
def Q(sql,a=()):
    with A.app.app_context(): return _q1(sql,a)
c = A.app.test_client()
def tok():
    c.get('/login')
    with c.session_transaction() as ss: return ss['csrf']
def post(u, **d):
    d['_csrf'] = tok(); return c.post(u, data=d, follow_redirects=True)
def as_(user, pw, role=None):
    c.get('/logout'); r = post('/login', username=user, password=pw); assert r.status_code == 200
assert post('/login', username='admin', password='admin123').status_code == 200
for u, r in (('acc', 'ACCOUNTANT'), ('chk', 'CHECKER'), ('apr', 'APPROVER'), ('view', 'VIEWER')): post('/users', username=u, name=u, password='pw1234', role=r)
as_('acc', 'pw1234')
post('/periods', label='Wk1', week='1', start='2026-01-05', end='2026-01-11', cal='G.C.')
post('/periods', label='Wk2', week='2', start='2026-01-12', end='2026-01-18', cal='G.C.')
post('/designs', name='Addis Design', cat='x', desc='', notes='')
post('/workers', name='Ambachew', phone='', start='', open='0', notes='')
post('/workers', name='Tewodros', phone='', start='', open='0', notes='')
post('/workers/1/price', design='1', price='900', eff='2026-01-01', notes='')
post('/workers/2/price', design='1', price='800', eff='2026-01-01', notes='')
post('/workers/1/price', design='1', price='1000', eff='2026-01-10', notes='raise')   # Wk1 starts 01-05 -> still 900; Wk2 -> 1000
post('/contract/add', p='1', w='1', d='1', qty='6', cust='Sam Customer', remark='x')
post('/contract/add', p='1', w='2', d='1', qty='6', cust='', remark='')
post('/contract/pay', p='1', pay_1='4000', pay_2='4800')
with A.app.app_context():
    A.app.test_request_context().push() if False else None
with A.app.test_request_context('/?p=1'):
    from flask import session
    b1 = A.bal(1, 1); b2 = A.bal(2, 1)
    assert (b1['earned'], b1['close']) == (5400, 1400), b1      # 6x900, 5400-4000
    assert (b2['earned'], b2['close']) == (4800, 0), b2         # 6x800
post('/contract/add', p='2', w='1', d='1', qty='2', cust='', remark='')
with A.app.test_request_context('/?p=2'):
    b = A.bal(1, 2); assert (b['prev'], b['earned'], b['payable']) == (1400, 2000, 3400), b   # new price 1000, carry 1400
    assert A.bal(1, 1)['earned'] == 5400                                                     # history unchanged
post('/employees', name='Wendesen', dept='', pos='', gross='10000', freq='WEEKLY', start='', notes='')
post('/employees', name='Eleven', dept='', pos='', gross='12000', freq='WEEKLY', start='', notes='')
post('/employees', name='Bi', dept='', pos='', gross='12000', freq='BIWEEKLY', start='', notes='')
post('/main/save', p='1', **{'1_tr': '2650'})
post('/overtime/add', p='1', emp='2', date='2026-01-06', type='NORMAL', end='20:30', hours='')      # 3 h after 17:30
post('/overtime/add', p='1', emp='2', date='2026-01-11', type='SUNDAY', end='', hours='')           # blank = standard 8 h
post('/overtime/add', p='1', emp='2', date='2026-01-07', type='NIGHT', end='', hours='2')
post('/overtime/add', p='1', emp='1', date='2026-01-08', type='NORMAL', end='', hours='2')         # manual hours path
post('/overtime/add', p='1', emp='1', date='2026-02-08', type='NORMAL', end='', hours='2')         # outside period -> refused
post('/overtime/add', p='1', emp='3', date='2026-01-08', type='NORMAL', end='', hours='2')         # biweekly not due -> refused
with A.app.test_request_context('/?p=1'):
    s = A.S(); r = {e['full_name']: A.calc(e, A.cur(), s) for e in A.q('select * from main_employees')}
    assert r['Wendesen']['ot_hours'] == 2 and r['Wendesen']['net_pay'] == round(2500 + 2650 + 2 * 10000 / 192 * 1.5, 2), r['Wendesen']
    assert r['Eleven']['ot_hours'] == 3 and r['Eleven']['ot_pay'] == 281.25 and r['Eleven']['sun_pay'] == 1000, r['Eleven']
    assert r['Eleven']['night_pay'] == 2 * 62.5 * 1.5
    assert r['Bi']['due'] == 0 and r['Bi']['net_pay'] == 0          # biweekly not due in week 1
    assert A.is_due('BIWEEKLY', 2)
post('/employees/2/salary', gross='15000', eff='2026-01-12', reason='raise')
with A.app.test_request_context('/'):
    e = A.q1('select * from main_employees where id=2'); assert A.gross_at(2, '2026-01-05') == 12000 and A.gross_at(2, '2026-01-12') == 15000

# adjustments (other earnings) + absence day-rate + csrf + password + backup
post('/adjustments', who='CONTRACT:2', amount='500', desc='fix', target='2', ref='1', reason='wk1 correction')
post('/adjustments', who='MAIN:1', amount='-100', desc='fine', target='2', ref='', reason='r')
with A.app.test_request_context('/?p=2'):
    assert A.bal(2, 2)['other'] == 500 and A.bal(2, 2)['payable'] == 500, A.bal(2, 2)
    e = A.q1('select * from main_employees where id=1'); P2 = A.q1('select * from payroll_periods where id=2'); assert A.calc(e, P2, A.S())['other_deductions'] == 100
with A.app.app_context(): A.ex("update settings set value='30' where key='absence_divisor'")
post('/main/save', p='1', **{'1_tr': '2650', '1_absd': '2'})
with A.app.test_request_context('/?p=1'):
    e = A.q1('select * from main_employees where id=1'); P1 = A.q1('select * from payroll_periods where id=1'); r = A.calc(e, P1, A.S())
    assert r['absence'] == round(10000 / 30 * 2, 2) and r['net_pay'] == round(2500 + 2650 + 156.25 - 10000 / 30 * 2, 2), r
with A.app.app_context(): A.ex("update settings set value='0' where key='absence_divisor'")
assert A.app.test_client().post('/designs', data={'name': 'X'}).status_code == 400          # no csrf -> rejected
post('/account', old='pw1234', new='newpass1', again='newpass1')
as_('acc', 'newpass1')
assert 'Markid Payroll' in c.get('/adjustments').get_data(as_text=True)
as_('admin', 'admin123'); assert c.get('/backup').status_code == 200; assert c.get('/export/statement/1').status_code == 200
post('/users/5/update', role='MANAGER', password='')
as_('acc', 'newpass1')
# workflow + permissions
post('/period/1/advance'); assert Q('select status s from payroll_periods where id=1')['s'] == 'CALCULATED'
assert 'permission' in post('/period/1/advance').get_data(as_text=True) or Q('select status s from payroll_periods where id=1')['s'] == 'CALCULATED'
as_('chk', 'pw1234'); post('/period/1/advance'); as_('apr', 'pw1234'); post('/period/1/advance')
as_('acc', 'newpass1'); post('/period/1/advance')
as_('apr', 'pw1234'); post('/period/1/advance'); assert Q('select status s from payroll_periods where id=1')['s'] == 'LOCKED'
as_('acc', 'newpass1'); n = Q('select count(*) c from contract_work_items')['c']; post('/contract/add', p='1', w='1', d='1', qty='5', cust='', remark='')
assert Q('select count(*) c from contract_work_items')['c'] == n            # locked: edit refused
for u in ('/', '/workers/1', '/employees/2', '/main?p=1', '/contract?p=1', '/audit', '/periods', '/settings', '/designs', '/workers', '/employees'):
    assert c.get(u).status_code == 200, u
for u in ('/export/main?p=1', '/export/contract?p=1'):
    r = c.get(u); assert r.status_code == 200 and r.data[:2] == b'PK', u
import io, openpyxl
wb = openpyxl.load_workbook(io.BytesIO(c.get('/export/contract?p=1').data)); txt = ' '.join(str(x.value) for row in wb['Payroll Report'].iter_rows() for x in row if x.value is not None)
assert 'Addis Design' in txt and 'Ambachew' in txt and 'Remaining balance' in txt and wb.sheetnames == ['Payroll Report', 'All Work Lines', 'By Design'], wb.sheetnames
wb = openpyxl.load_workbook(io.BytesIO(c.get('/export/main?p=1').data)); assert wb.sheetnames == ['Payroll Report', 'Pay Breakdown', 'Overtime Entries']
ot = [[x.value for x in row] for row in wb['Overtime Entries'].iter_rows(min_row=5)]; assert any(r[3] == 'Sunday / weekend' and r[6] == 8 for r in ot), ot
as_('view', 'pw1234'); assert c.post('/designs', data={'name': 'Z', '_csrf': tok()}).status_code == 403
print('ALL OK')
