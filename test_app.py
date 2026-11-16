import os, tempfile, importlib, json, io
os.environ['MAXUP_DATA_DIR']=tempfile.mkdtemp(prefix='maxup-tests-')
import app as m
from werkzeug.security import generate_password_hash
import pytest

@pytest.fixture()
def clients():
 with m.db() as c:
  c.execute('DELETE FROM imports');c.execute('DELETE FROM approvals');c.execute('DELETE FROM sales');c.execute('DELETE FROM lots')
  for email,role in [('admin@test.local','admin'),('sales@test.local','sales')]:
   c.execute('INSERT OR IGNORE INTO users(email,password,role) VALUES(?,?,?)',(email,generate_password_hash('Very-long-password!'),role))
 a=m.app.test_client();s=m.app.test_client()
 ta=a.post('/api/login',json={'email':'admin@test.local','password':'Very-long-password!'}).json['csrf']
 ts=s.post('/api/login',json={'email':'sales@test.local','password':'Very-long-password!'}).json['csrf']
 return a,s,{'X-CSRF-Token':ta},{'X-CSRF-Token':ts}

def workbook(sku='TEST-A',price=120000):
 from openpyxl import Workbook
 w=Workbook();s=w.active;s.append(m.COLUMNS);s.append([sku,'Test Product','Forklift','2.5T',price,price,price,price,2,4,110000,None,'2026-01-01']);o=io.BytesIO();w.save(o);o.seek(0);return o

def upload(a,h,price=120000):
 return a.post('/api/import/preview',data={'file':(workbook(price=price),'prices.xlsx')},headers=h)
def seed(a,h):
 r=upload(a,h); assert r.status_code==200,r.json
 r=a.post('/api/import/confirm',json={'id':r.json['id'],'reason':'test import'},headers=h);assert r.status_code==200,r.json

def test_access_and_cost_isolation(clients):
 a,s,ha,hs=clients;seed(a,ha)
 r=a.post('/api/lots',json=dict(sku='TEST-A',ref='LOT-1',currency='USD',po=2500,fx=36,factor=1.1,actual=100000,quantity=3),headers=ha);assert r.status_code==200,r.json
 body=s.get('/api/products?date=2026-11-16').json
 assert body and not any(k in body[0] for k in ['lots','cost','cost_status','profit','actual','estimated'])
 for endpoint in ['analytics','audit','import/template']:
  assert s.get('/api/'+endpoint).status_code==403
 assert s.post('/api/lots',json={},headers=hs).status_code==403
 assert a.post('/api/prices',json={},headers={}).status_code==403
 assert m.app.test_client().get('/api/products').status_code==401

def test_effective_dates_and_immutable_history(clients):
 a,s,ha,hs=clients;seed(a,ha)
 x=dict(sku='TEST-A',p1=150000,p2=150000,p3=150000,dealer=150000,floor=110000,market=None,discount34=2,discount5=4,effective='2099-01-01',reason='future')
 assert a.post('/api/prices',json=x,headers=ha).status_code==200
 assert s.get('/api/products?date=2026-11-16').json[0]['p1']==120000
 assert s.get('/api/products?date=2099-01-01').json[0]['p1']==150000
 with pytest.raises(Exception):
  with m.db() as c:c.execute('UPDATE prices SET p1=0')

def test_import_repeat_backup_and_stale_preview(clients):
 a,s,ha,hs=clients;seed(a,ha);p=upload(a,ha);iid=p.json['id']
 assert p.json['rows'][0]['status']=='unchanged'
 assert a.post('/api/import/confirm',json={'id':iid,'reason':'unchanged'},headers=ha).status_code==200
 assert a.post('/api/import/confirm',json={'id':iid,'reason':'repeat'},headers=ha).status_code==400
 assert list((m.DATA/'backups').glob('*.sqlite3'))
 preview=upload(a,ha,125000).json
 x=dict(sku='TEST-A',p1=130000,p2=130000,p3=130000,dealer=130000,floor=110000,discount34=2,discount5=4,effective='2026-02-01',reason='concurrent change')
 assert a.post('/api/prices',json=x,headers=ha).status_code==200
 assert a.post('/api/import/confirm',json={'id':preview['id'],'reason':'stale'},headers=ha).status_code==409

def test_sales_inventory_and_realized_profit(clients):
 a,s,ha,hs=clients;seed(a,ha)
 data=dict(sku='TEST-A',ref='LOT-A',currency='USD',po=2500,fx=36,factor=1.1,actual=100000,quantity=2)
 assert a.post('/api/lots',json=data,headers=ha).status_code==200
 lot=a.get('/api/products').json[0]['lots'][0]
 assert lot['estimated']==9900000
 sale=dict(invoice='INV-001',lot=lot['id'],quantity=1,unit_price=120000,date='2026-11-16')
 assert a.post('/api/sales',json=sale,headers=ha).status_code==200
 x=a.get('/api/analytics').json
 assert x['inventory_actual']==100000
 assert x['monthly'][0]['revenue']==120000 and x['monthly'][0]['profit']==20000
 assert a.post('/api/sales',json=sale,headers=ha).status_code==400
 assert a.get('/api/products').json[0]['lots'][0]['quantity']==1
 sale['quantity']=2
 assert a.post('/api/sales',json=sale,headers=ha).status_code==400
 assert a.get('/api/products').json[0]['lots'][0]['quantity']==1
 assert a.post('/api/lots/'+str(lot['id'])+'/actual',json={'actual':1,'reason':'tamper'},headers=ha).status_code==400

def test_import_validation_and_pending_cost(clients):
 a,s,ha,hs=clients;seed(a,ha)
 assert a.post('/api/lots',json=dict(sku='TEST-A',ref='PENDING',currency='RMB',po=2000,fx=5,factor=1.15,quantity=2),headers=ha).status_code==200
 lot=a.get('/api/products').json[0]['lots'][0]
 assert a.post('/api/sales',json=dict(invoice='INV-2',lot=lot['id'],quantity=1,unit_price=120000,date='2026-11-16'),headers=ha).status_code==400
 bad=a.post('/api/import/preview',data={'file':(io.BytesIO(b'invalid'),'wrong.csv')},headers=ha)
 assert bad.status_code==400
