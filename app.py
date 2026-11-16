"""MAXUP Pricing — standalone reference implementation. Monetary values: satang integers."""
import os, json, sqlite3, secrets, hashlib, io, zipfile
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path
from functools import wraps
from flask import Flask, request, session, jsonify, render_template, send_file, abort
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import load_workbook, Workbook

ROOT=Path(__file__).parent
app=Flask(__name__)
app.config.update(SECRET_KEY=os.environ.get('MAXUP_SECRET_KEY') or secrets.token_hex(32), MAX_CONTENT_LENGTH=5*1024*1024,
 SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE='Strict', SESSION_COOKIE_SECURE=os.environ.get('MAXUP_HTTPS')=='1', PERMANENT_SESSION_LIFETIME=timedelta(hours=8))
DATA=Path(os.environ.get('MAXUP_DATA_DIR',ROOT/'instance')); DATA.mkdir(parents=True,exist_ok=True)
DB=DATA/'maxup.sqlite3'
def now(): return datetime.now(ZoneInfo('Asia/Bangkok')).isoformat(timespec='seconds')
def db():
 c=sqlite3.connect(DB,timeout=20); c.row_factory=sqlite3.Row; c.execute('PRAGMA foreign_keys=ON'); return c
def money(v):
 d=Decimal(str(v))
 if not d.is_finite() or d<0 or d>Decimal('999999999'): raise ValueError('จำนวนเงินไม่ถูกต้อง')
 return int((d*100).quantize(Decimal('1'),rounding=ROUND_HALF_UP))
def required(s):
 s=str(s or '').strip()
 if not s or len(s)>2000: raise ValueError('กรุณากรอกข้อมูลให้ครบ')
 return s

def init():
 with db() as c:
  c.executescript('''
 CREATE TABLE IF NOT EXISTS users(id INTEGER PRIMARY KEY,email TEXT UNIQUE NOT NULL,password TEXT NOT NULL,role TEXT NOT NULL CHECK(role IN ('sales','admin','executive')),active INTEGER NOT NULL DEFAULT 1);
 CREATE TABLE IF NOT EXISTS attempts(key TEXT PRIMARY KEY,n INTEGER NOT NULL,until TEXT NOT NULL);
 CREATE TABLE IF NOT EXISTS products(sku TEXT PRIMARY KEY,name TEXT NOT NULL,category TEXT NOT NULL,spec TEXT NOT NULL);
 CREATE TABLE IF NOT EXISTS prices(id INTEGER PRIMARY KEY,sku TEXT NOT NULL REFERENCES products(sku),effective TEXT NOT NULL,p1 INTEGER NOT NULL,p2 INTEGER NOT NULL,p3 INTEGER NOT NULL,dealer INTEGER NOT NULL,discount34 REAL NOT NULL,discount5 REAL NOT NULL,floor INTEGER NOT NULL,market INTEGER,actor INTEGER NOT NULL REFERENCES users(id),reason TEXT NOT NULL,created TEXT NOT NULL);
 CREATE INDEX IF NOT EXISTS price_date ON prices(sku,effective,id);
 CREATE TABLE IF NOT EXISTS lots(id INTEGER PRIMARY KEY,sku TEXT NOT NULL REFERENCES products(sku),ref TEXT UNIQUE NOT NULL,currency TEXT NOT NULL,po_minor INTEGER NOT NULL,fx TEXT NOT NULL,factor TEXT NOT NULL,estimated INTEGER NOT NULL,actual INTEGER,quantity INTEGER NOT NULL CHECK(quantity>=0),created TEXT NOT NULL);
 CREATE TABLE IF NOT EXISTS sales(id INTEGER PRIMARY KEY,invoice TEXT NOT NULL,lot INTEGER NOT NULL REFERENCES lots(id),quantity INTEGER NOT NULL CHECK(quantity>0),unit_price INTEGER NOT NULL,cost INTEGER NOT NULL,date TEXT NOT NULL,actor INTEGER NOT NULL REFERENCES users(id));
 CREATE UNIQUE INDEX IF NOT EXISTS sales_document_lot ON sales(invoice,lot);
 CREATE TABLE IF NOT EXISTS approvals(id INTEGER PRIMARY KEY,sku TEXT NOT NULL REFERENCES products(sku),price INTEGER NOT NULL,quantity INTEGER NOT NULL,reason TEXT NOT NULL,actor INTEGER NOT NULL REFERENCES users(id),status TEXT NOT NULL DEFAULT 'pending',reviewer INTEGER REFERENCES users(id),created TEXT NOT NULL);
 CREATE TABLE IF NOT EXISTS audit(id INTEGER PRIMARY KEY,actor INTEGER NOT NULL,action TEXT NOT NULL,payload TEXT NOT NULL,created TEXT NOT NULL);
 CREATE TABLE IF NOT EXISTS imports(id TEXT PRIMARY KEY,actor INTEGER NOT NULL,base TEXT NOT NULL,payload TEXT NOT NULL,created TEXT NOT NULL,used INTEGER NOT NULL DEFAULT 0);
 CREATE TRIGGER IF NOT EXISTS prices_no_update BEFORE UPDATE ON prices BEGIN SELECT RAISE(ABORT,'immutable'); END;
 CREATE TRIGGER IF NOT EXISTS prices_no_delete BEFORE DELETE ON prices BEGIN SELECT RAISE(ABORT,'immutable'); END;
 CREATE TRIGGER IF NOT EXISTS audit_no_update BEFORE UPDATE ON audit BEGIN SELECT RAISE(ABORT,'immutable'); END;
 CREATE TRIGGER IF NOT EXISTS audit_no_delete BEFORE DELETE ON audit BEGIN SELECT RAISE(ABORT,'immutable'); END;
 ''')
init()

def auth(admin=False):
 def deco(fn):
  @wraps(fn)
  def wrapped(*a,**kw):
   with db() as c: u=c.execute('SELECT id,email,role FROM users WHERE id=? AND active=1',(session.get('uid'),)).fetchone()
   if not u: abort(401)
   if admin and u['role'] not in ('admin','executive'): abort(403)
   request.user=dict(u)
   if request.method!='GET' and not secrets.compare_digest(request.headers.get('X-CSRF-Token',''),session.get('csrf','INVALID')): abort(403)
   return fn(*a,**kw)
  return wrapped
 return deco

def audit(c,action,payload): c.execute('INSERT INTO audit(actor,action,payload,created) VALUES(?,?,?,?)',(request.user['id'],action,json.dumps(payload,ensure_ascii=False),now()))
def latest(c,sku,date): return c.execute('SELECT * FROM prices WHERE sku=? AND effective<=? ORDER BY effective DESC,id DESC LIMIT 1',(sku,date)).fetchone()
def fingerprint(c):
 content=[list(r) for r in c.execute('SELECT * FROM products ORDER BY sku')]+[list(r) for r in c.execute('SELECT * FROM prices ORDER BY id')]
 return hashlib.sha256(json.dumps(content).encode()).hexdigest()
def validate_price(x):
 p={k:money(x[k]) for k in ['p1','p2','p3','dealer','floor']}
 p['market']=money(x['market']) if x.get('market') not in (None,'') else None
 for k in ['discount34','discount5']:
  v=float(x.get(k,0))
  if not 0<=v<=100: raise ValueError('ส่วนลดต้องอยู่ระหว่าง 0–100%')
  p[k]=v
 if p['floor']>min(p[k] for k in ['p1','p2','p3','dealer']): raise ValueError('Floor Price ต้องไม่สูงกว่าระดับราคาขาย')
 p['effective']=datetime.fromisoformat(required(x['effective'])).date().isoformat()
 return p

def insert_price(c,sku,x,reason):
 p=validate_price(x)
 c.execute('INSERT INTO prices(sku,effective,p1,p2,p3,dealer,discount34,discount5,floor,market,actor,reason,created) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',
 (sku,p['effective'],p['p1'],p['p2'],p['p3'],p['dealer'],p['discount34'],p['discount5'],p['floor'],p['market'],request.user['id'],required(reason),now()))

def present(p):
 p=dict(p)
 for k in ['p1','p2','p3','dealer','floor','market']:
  if p.get(k) is not None: p[k]=p[k]/100
 return p

@app.after_request
def headers(r):
 r.headers['Cache-Control']='no-store'; r.headers['X-Content-Type-Options']='nosniff'; r.headers['X-Frame-Options']='DENY'
 r.headers['Content-Security-Policy']="default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; base-uri 'none'; frame-ancestors 'none'; form-action 'self'"
 return r
@app.errorhandler(Exception)
def errors(e):
 from werkzeug.exceptions import HTTPException
 if isinstance(e,HTTPException): return jsonify(error={401:'กรุณาเข้าสู่ระบบ',403:'ไม่มีสิทธิ์ดำเนินการ',404:'ไม่พบรายการ'}.get(e.code,e.description)),e.code
 if isinstance(e,(ValueError,KeyError,TypeError,InvalidOperation,zipfile.BadZipFile,sqlite3.IntegrityError)): return jsonify(error=str(e)),400
 app.logger.exception('Request failed'); return jsonify(error='ระบบขัดข้อง กรุณาตรวจสอบ server log'),500
@app.get('/')
def home(): return render_template('index.html')
@app.post('/api/login')
def login():
 x=request.get_json(); email=str(x.get('email','')).lower().strip(); key=hashlib.sha256((request.remote_addr+'|'+email).encode()).hexdigest()
 with db() as c:
  a=c.execute('SELECT * FROM attempts WHERE key=?',(key,)).fetchone()
  if a and a['n']>=8 and a['until']>now(): return jsonify(error='ลองใหม่ในอีก 15 นาที'),429
  u=c.execute('SELECT * FROM users WHERE email=? AND active=1',(email,)).fetchone()
  if not u or not check_password_hash(u['password'],x.get('password','')):
   n=a['n']+1 if a and a['until']>now() else 1
   c.execute('INSERT OR REPLACE INTO attempts VALUES(?,?,?)',(key,n,(datetime.now(ZoneInfo('Asia/Bangkok'))+timedelta(minutes=15)).isoformat(timespec='seconds')))
   return jsonify(error='อีเมลหรือรหัสผ่านไม่ถูกต้อง'),401
  c.execute('DELETE FROM attempts WHERE key=?',(key,))
 session.clear(); session.permanent=True; session['uid']=u['id']; session['csrf']=secrets.token_hex(32)
 return jsonify(email=u['email'],role=u['role'],csrf=session['csrf'])
@app.get('/api/me')
@auth()
def me(): return jsonify(**request.user,csrf=session['csrf'])
@app.post('/api/logout')
@auth()
def logout(): session.clear(); return jsonify(ok=True)
@app.get('/api/products')
@auth()
def products():
 date=request.args.get('date',now()[:10]); datetime.fromisoformat(date); result=[]; admin=request.user['role']!='sales'
 with db() as c:
  for prod in c.execute('SELECT * FROM products ORDER BY category,sku'):
   p=latest(c,prod['sku'],date)
   if not p: continue
   item=dict(prod); item.update(present({k:p[k] for k in ['p1','p2','p3','dealer','discount34','discount5','floor','market','effective']}))
   if admin:
    lots=[dict(r) for r in c.execute('SELECT * FROM lots WHERE sku=? ORDER BY id DESC',(prod['sku'],))]
    item['lots']=lots
    lot=lots[0] if lots else None
    item['cost']=((lot['actual'] if lot['actual'] is not None else lot['estimated'])/100) if lot else None
    item['cost_status']='actual' if lot and lot['actual'] is not None else 'estimated' if lot else 'missing'
   result.append(item)
 return jsonify(result)
@app.post('/api/prices')
@auth(True)
def prices():
 x=request.get_json()
 with db() as c:
  insert_price(c,required(x['sku']),x,x['reason']); audit(c,'price_change',x)
 return jsonify(ok=True)
@app.get('/api/history')
@auth()
def history():
 with db() as c:
  rows=c.execute('SELECT p.*,u.email FROM prices p JOIN users u ON p.actor=u.id WHERE sku=? ORDER BY effective DESC,p.id DESC',(request.args.get('sku'),)).fetchall()
 return jsonify([present({k:r[k] for k in ['effective','p1','p2','p3','dealer','floor','market','discount34','discount5','email','reason','created']}) for r in rows])
@app.post('/api/lots')
@auth(True)
def lots():
 x=request.get_json(); currency=x['currency']; qty=int(x['quantity']); fx=Decimal(str(x['fx'])); factor=Decimal(str(x['factor']))
 if currency not in ('USD','RMB','THB') or qty<0 or not fx.is_finite() or not 0<fx<=10000 or not factor.is_finite() or not 1<=factor<=100: raise ValueError('ตรวจสอบสกุลเงิน เรต และตัวคูณนำเข้า')
 po=money(x['po']); estimated=int((Decimal(po)*fx*factor).quantize(Decimal('1'),rounding=ROUND_HALF_UP)); actual=money(x['actual']) if x.get('actual') not in ('',None) else None
 with db() as c:
  c.execute('INSERT INTO lots(sku,ref,currency,po_minor,fx,factor,estimated,actual,quantity,created) VALUES(?,?,?,?,?,?,?,?,?,?)',(x['sku'],required(x['ref']),currency,po,str(fx),str(factor),estimated,actual,qty,now())); audit(c,'create_lot',x)
 return jsonify(ok=True)
@app.post('/api/lots/<int:lot_id>/actual')
@auth(True)
def actual(lot_id):
 x=request.get_json(); required(x['reason'])
 with db() as c:
  if not c.execute('SELECT 1 FROM lots WHERE id=?',(lot_id,)).fetchone(): abort(404)
  if c.execute('SELECT 1 FROM sales WHERE lot=?',(lot_id,)).fetchone(): raise ValueError('ลอตนี้มีการขายแล้ว กรุณาใช้กระบวนการปรับปรุงบัญชีแทนการแก้ต้นทุนย้อนหลัง')
  c.execute('UPDATE lots SET actual=? WHERE id=?',(money(x['actual']),lot_id)); audit(c,'actual_cost',dict(lot=lot_id,**x))
 return jsonify(ok=True)
@app.post('/api/sales')
@auth(True)
def sale():
 x=request.get_json(); qty=int(x['quantity']); unit=money(x['unit_price']); date=datetime.fromisoformat(x['date']).date().isoformat()
 if qty<=0: raise ValueError('จำนวนต้องมากกว่า 0')
 with db() as c:
  c.execute('BEGIN IMMEDIATE'); lot=c.execute('SELECT * FROM lots WHERE id=?',(int(x['lot']),)).fetchone()
  if not lot or lot['quantity']<qty: raise ValueError('สต็อกไม่เพียงพอ')
  if lot['actual'] is None: raise ValueError('ต้องบันทึกต้นทุนจริงก่อนรับรู้กำไร')
  c.execute('UPDATE lots SET quantity=quantity-? WHERE id=?',(qty,lot['id']))
  c.execute('INSERT INTO sales(invoice,lot,quantity,unit_price,cost,date,actor) VALUES(?,?,?,?,?,?,?)',(required(x['invoice']),lot['id'],qty,unit,lot['actual'],date,request.user['id'])); audit(c,'sale',x)
 return jsonify(ok=True)
@app.get('/api/analytics')
@auth(True)
def analytics():
 with db() as c:
  monthly=[dict(r) for r in c.execute("SELECT substr(date,1,7) month,SUM(quantity*unit_price)/100.0 revenue,SUM(quantity*(unit_price-cost))/100.0 profit FROM sales GROUP BY month ORDER BY month DESC")]
  actual=c.execute('SELECT COALESCE(SUM(quantity*actual),0)/100.0 FROM lots WHERE actual IS NOT NULL').fetchone()[0]
  estimate=c.execute('SELECT COALESCE(SUM(quantity*estimated),0)/100.0 FROM lots WHERE actual IS NULL').fetchone()[0]
 return jsonify(monthly=monthly,inventory_actual=actual,inventory_estimated=estimate)
@app.route('/api/approvals',methods=['GET','POST'])
@auth()
def approvals():
 with db() as c:
  if request.method=='POST':
   x=request.get_json(); qty=int(x['quantity'])
   if qty<=0: raise ValueError('จำนวนต้องมากกว่า 0')
   c.execute('INSERT INTO approvals(sku,price,quantity,reason,actor,created) VALUES(?,?,?,?,?,?)',(x['sku'],money(x['price']),qty,required(x['reason']),request.user['id'],now())); audit(c,'approval_request',x); return jsonify(ok=True)
  query='SELECT a.*,u.email FROM approvals a JOIN users u ON a.actor=u.id'; params=()
  if request.user['role']=='sales': query+=' WHERE a.actor=?'; params=(request.user['id'],)
  return jsonify([dict(r) for r in c.execute(query+' ORDER BY a.id DESC',params)])
@app.post('/api/approvals/<int:aid>')
@auth(True)
def decide(aid):
 x=request.get_json(); status=x['status']
 if status not in ('approved','rejected'): raise ValueError('สถานะไม่ถูกต้อง')
 with db() as c:
  r=c.execute('UPDATE approvals SET status=?,reviewer=? WHERE id=? AND status=?',(status,request.user['id'],aid,'pending'))
  if not r.rowcount: raise ValueError('รายการได้รับการพิจารณาแล้ว')
  audit(c,'approval_decision',dict(id=aid,status=status))
 return jsonify(ok=True)

COLUMNS=['sku','name','category','spec','p1','p2','p3','dealer','discount34','discount5','floor','market','effective']
@app.get('/api/import/template')
@auth(True)
def template():
 w=Workbook(); s=w.active; s.title='Prices'; s.append(COLUMNS); s.append(['DEMO-001','สินค้าตัวอย่าง','Forklift','2.5 ตัน / 3 เมตร',490000,475000,460000,445000,2,4,430000,None,now()[:10]])
 out=io.BytesIO(); w.save(out); out.seek(0); return send_file(out,download_name='MAXUP_Import_Template.xlsx',as_attachment=True)
@app.post('/api/import/preview')
@auth(True)
def preview():
 f=request.files.get('file')
 if not f or not f.filename.lower().endswith('.xlsx'): raise ValueError('รองรับเฉพาะ .xlsx')
 raw=f.read()
 with zipfile.ZipFile(io.BytesIO(raw)) as z:
  if sum(i.file_size for i in z.infolist())>20*1024*1024: raise ValueError('ไฟล์ใหญ่เกินกำหนด')
 w=load_workbook(io.BytesIO(raw),read_only=True,data_only=False); s=w.active
 if s.max_row>5001: raise ValueError('รองรับสูงสุด 5,000 แถว')
 rows=s.iter_rows(values_only=True)
 if list(next(rows))!=COLUMNS: raise ValueError('หัวตารางไม่ตรงกับ Template')
 result=[]; seen=set()
 with db() as c:
  for vals in rows:
   if all(v is None for v in vals): continue
   x=dict(zip(COLUMNS,vals))
   for k,v in x.items():
    if isinstance(v,str) and v.startswith('='): raise ValueError('ไม่รับสูตร Excel โปรดวางเป็นค่า')
   for k in ['sku','name','category','spec']: x[k]=required(x[k])
   if x['sku'] in seen: raise ValueError('SKU ซ้ำในไฟล์: '+x['sku'])
   seen.add(x['sku'])
   if isinstance(x['effective'],datetime): x['effective']=x['effective'].date().isoformat()
   p=validate_price(x); prod=c.execute('SELECT * FROM products WHERE sku=?',(x['sku'],)).fetchone(); price=latest(c,x['sku'],p['effective']); old=dict(prod) if prod else {}
   if price: old.update(present({k:price[k] for k in COLUMNS if k in price.keys()}))
   changes={k:{'old':old.get(k),'new':x[k]} for k in COLUMNS if old.get(k)!=x[k]}
   result.append(dict(status='new' if not prod else 'modified' if changes else 'unchanged',data=x,changes=changes))
  iid=secrets.token_hex(16); c.execute('INSERT INTO imports VALUES(?,?,?,?,?,0)',(iid,request.user['id'],fingerprint(c),json.dumps(result,ensure_ascii=False),now()))
 return jsonify(id=iid,rows=result)
@app.post('/api/import/confirm')
@auth(True)
def confirm():
 x=request.get_json(); reason=required(x['reason'])
 with db() as c:
  c.execute('BEGIN IMMEDIATE'); imp=c.execute('SELECT * FROM imports WHERE id=? AND actor=? AND used=0',(x['id'],request.user['id'])).fetchone()
  if not imp: raise ValueError('ไม่พบรายการ หรือ Import นี้ถูกใช้แล้ว')
  if fingerprint(c)!=imp['base']: return jsonify(error='ข้อมูลเปลี่ยนหลัง Preview กรุณาอัปโหลดเพื่อตรวจสอบอีกครั้ง'),409
  if datetime.fromisoformat(imp['created'])+timedelta(minutes=30)<datetime.now(ZoneInfo('Asia/Bangkok')): raise ValueError('Preview หมดอายุ กรุณาอัปโหลดใหม่')
  backup=DATA/'backups'; backup.mkdir(exist_ok=True)
  # Online SQLite backup from a separate read connection, before any mutation.
  with db() as src, sqlite3.connect(backup/(now()[:10]+'-'+secrets.token_hex(8)+'.sqlite3')) as dst: src.backup(dst)
  for row in json.loads(imp['payload']):
   if row['status']=='unchanged': continue
   p=row['data']; c.execute('INSERT INTO products VALUES(?,?,?,?) ON CONFLICT(sku) DO UPDATE SET name=excluded.name,category=excluded.category,spec=excluded.spec',(p['sku'],p['name'],p['category'],p['spec']))
   insert_price(c,p['sku'],p,reason)
  c.execute('UPDATE imports SET used=1 WHERE id=?',(x['id'],)); audit(c,'import',dict(id=x['id'],reason=reason))
 return jsonify(ok=True)
@app.get('/api/audit')
@auth(True)
def logs():
 with db() as c: return jsonify([dict(r) for r in c.execute('SELECT a.*,u.email FROM audit a JOIN users u ON a.actor=u.id ORDER BY a.id DESC LIMIT 200')])

if __name__=='__main__':
 import argparse, getpass
 p=argparse.ArgumentParser(); p.add_argument('command',choices=['run','user','demo']); args=p.parse_args()
 if args.command=='user':
  email=input('Email: ').strip().lower(); role=input('Role (sales/admin/executive): ').strip(); password=getpass.getpass('Password (12+ characters): ')
  if role not in ('sales','admin','executive') or len(password)<12: raise SystemExit('Invalid role or short password')
  with db() as c: c.execute('INSERT INTO users(email,password,role) VALUES(?,?,?)',(email,generate_password_hash(password),role))
  print('User created')
 elif args.command=='demo':
  with db() as c:
   u=c.execute("SELECT id FROM users WHERE role IN ('admin','executive') LIMIT 1").fetchone()
   if not u: raise SystemExit('Create an admin user first')
   for i,(sku,name,cat,spec,price,cost) in enumerate([
    ('DEMO-FB25','Electric Forklift','Forklift','2.5 ตัน • เสา 3 เมตร • Lithium 80V',490000,365000),
    ('DEMO-FB30','Electric Forklift','Forklift','3.0 ตัน • เสา 4.5 เมตร • Lithium',560000,420000),
    ('DEMO-PS15','Electric Stacker','Stacker','1.5 ตัน • ยกสูง 3 เมตร',99000,73000),
    ('DEMO-PS20','Electric Stacker','Stacker','2.0 ตัน • ยกสูง 3.5 เมตร',159000,119000),
    ('DEMO-CBY30','Hand Pallet Truck','Pallet Truck','3.0 ตัน • งา 685 × 1,220 มม.',8900,6100),
    ('DEMO-XL12','Scissor Lift','Scissor Lift','ตัวอย่างสเปก • โปรดตรวจสอบ Catalog',330000,260000)]):
    if c.execute('SELECT 1 FROM products WHERE sku=?',(sku,)).fetchone(): continue
    c.execute('INSERT INTO products VALUES(?,?,?,?)',(sku,name,cat,spec))
    c.execute('INSERT INTO prices(sku,effective,p1,p2,p3,dealer,discount34,discount5,floor,market,actor,reason,created) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)',(sku,now()[:10],price*100,int(price*.97)*100,int(price*.94)*100,int(price*.9)*100,2,4,int(price*.88)*100,None,u['id'],'DEMO ONLY — ไม่ใช่ราคาจริง',now()))
    c.execute('INSERT INTO lots(sku,ref,currency,po_minor,fx,factor,estimated,actual,quantity,created) VALUES(?,?,?,?,?,?,?,?,?,?)',(sku,'DEMO-LOT-'+str(i),'THB',cost*100,'1','1',cost*100,cost*100 if i%2==0 else None,3+i,now()))
  print('Demo records inserted — NOT official MAXUP prices')
 else: app.run(host='127.0.0.1',port=8000,debug=False)
