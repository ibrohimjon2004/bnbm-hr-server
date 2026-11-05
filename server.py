"""BNBM HR Telegram Mini App API. Run with uvicorn server:app --host 0.0.0.0 --port 8000"""
import os, json, hmac, hashlib, time, re, sqlite3, logging, asyncio
from datetime import datetime, timezone
from pathlib import Path
from io import BytesIO
from urllib.parse import parse_qsl
from contextlib import asynccontextmanager
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException, Header, UploadFile, File, Form, Query
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field
from openpyxl import Workbook
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.utils import ImageReader
from xml.sax.saxutils import escape
import httpx

load_dotenv()
ROOT=Path(__file__).resolve().parent
DB_PATH=os.getenv('DB_PATH',str(ROOT/'bnbm_hr.sqlite3'))
TOKEN=os.getenv('BOT_TOKEN','').strip()
ADMIN_ID=int(os.getenv('ADMIN_ID','0') or '0')
HR_IDS={int(x.strip()) for x in os.getenv('HR_IDS','').split(',') if x.strip().isdigit()}
MAX_PHOTO=5*1024*1024
FIELDS=['gender','first_name','last_name','birth_date','marital','address','phone','russian','student','shifts','experience','salary']
LABELS_ZH={'gender':'性别','first_name':'名','last_name':'姓','birth_date':'出生日期','marital':'婚姻状况','address':'居住地址','phone':'电话','russian':'俄语水平','student':'是否在校学生','shifts':'是否接受轮班','experience':'工作经历','salary':'期望月薪','crane_type':'起重机操作经验','certificate':'操作资格证'}
TRANSLATE={'Erkak':'男','Ayol':'女','Uylangan / Turmushga chiqqan':'已婚','Bo‘ydoq / Turmushga chiqmagan':'未婚','Ajrashgan':'离异','Beva':'丧偶','Umuman bilmayman':'不会','Boshlang‘ich':'初级','O‘rta':'中级','Yaxshi':'良好','Erkin muloqot qila olaman':'流利','Ha':'是','Yo‘q':'否','Olishni rejalashtirganman':'计划考取'}

def db():
    c=sqlite3.connect(DB_PATH,timeout=20)
    c.row_factory=sqlite3.Row
    return c

def init_db():
    with db() as c:
        c.executescript('''
        CREATE TABLE IF NOT EXISTS vacancies (id INTEGER PRIMARY KEY AUTOINCREMENT,department TEXT NOT NULL,title_uz TEXT NOT NULL,title_zh TEXT NOT NULL,description_uz TEXT DEFAULT '',description_zh TEXT DEFAULT '',active INTEGER DEFAULT 1);
        CREATE TABLE IF NOT EXISTS extra_questions (id INTEGER PRIMARY KEY AUTOINCREMENT,vacancy_id INTEGER NOT NULL,question_uz TEXT NOT NULL,question_zh TEXT NOT NULL,FOREIGN KEY(vacancy_id) REFERENCES vacancies(id));
        CREATE TABLE IF NOT EXISTS applications (id INTEGER PRIMARY KEY AUTOINCREMENT,telegram_id INTEGER NOT NULL,language TEXT NOT NULL,vacancy_id INTEGER NOT NULL,answers TEXT DEFAULT '',photo_file_id TEXT DEFAULT '',status TEXT DEFAULT 'Yangi',created_at TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS web_photos (application_id INTEGER PRIMARY KEY,photo BLOB NOT NULL,mime TEXT NOT NULL);
        ''')
        if c.execute('SELECT count(*) FROM vacancies').fetchone()[0]==0:
            c.executemany('INSERT INTO vacancies (department,title_uz,title_zh,description_uz,description_zh) VALUES (?,?,?,?,?)',[
                ('Savdo','Savdo mutaxassisi','销售专员','Mijozlar bilan ishlash va savdo faoliyati. Batafsil ma’lumot HR bo‘limidan olinadi.','负责客户沟通与销售工作。详情请咨询人力资源部。'),
                ('Ishlab chiqarish','Kran haydovchisi','起重机司机','Yuklarni ko‘tarish va tushirish, uskunani xavfsiz boshqarish. 12 soatlik smena, 3 mahal ovqat.','负责货物吊装、卸载，安全操作设备。12小时轮班，提供每日三餐。')])

@asynccontextmanager
async def lifespan(app):
    init_db()
    yield

app=FastAPI(title='BNBM HR API',lifespan=lifespan)

def user_from_init(init_data):
    if not TOKEN: raise HTTPException(503,'BOT_TOKEN sozlanmagan')
    if not init_data: raise HTTPException(401,'Telegram ilovasi ichidan oching')
    try:
        data=dict(parse_qsl(init_data,keep_blank_values=True,strict_parsing=True))
        received=data.pop('hash')
        auth_date=int(data['auth_date'])
        if abs(time.time()-auth_date)>3600: raise ValueError('expired')
        check='\n'.join(f'{k}={v}' for k,v in sorted(data.items()))
        secret=hmac.new(b'WebAppData',TOKEN.encode(),hashlib.sha256).digest()
        expected=hmac.new(secret,check.encode(),hashlib.sha256).hexdigest()
        if not hmac.compare_digest(expected,received): raise ValueError('signature')
        user=json.loads(data['user'])
        uid=int(user['id'])
        if uid<=0: raise ValueError('user')
        return uid
    except (KeyError,ValueError,TypeError,json.JSONDecodeError):
        raise HTTPException(401,'Telegram autentifikatsiyasi noto‘g‘ri yoki muddati tugagan')

def auth(x_telegram_init_data: str=Header(default='')):
    return user_from_init(x_telegram_init_data)

def admin(x_telegram_init_data: str=Header(default='')):
    uid=user_from_init(x_telegram_init_data)
    if uid not in ({ADMIN_ID}|HR_IDS): raise HTTPException(403,'HR huquqi yo‘q')
    return uid

@app.get('/api/me')
def me(uid:int=__import__('fastapi').Depends(auth)):
    return {'id':uid,'is_hr':uid in ({ADMIN_ID}|HR_IDS)}

@app.get('/api/vacancies')
def vacancies(uid:int=__import__('fastapi').Depends(auth)):
    with db() as c:
        vs=[dict(x) for x in c.execute('SELECT * FROM vacancies WHERE active=1 ORDER BY id DESC')]
        for v in vs:
            v['questions']=[dict(x) for x in c.execute('SELECT id,question_uz,question_zh FROM extra_questions WHERE vacancy_id=? ORDER BY id',(v['id'],))]
    return vs

@app.get('/api/admin/vacancies')
def all_vacancies(uid:int=__import__('fastapi').Depends(admin)):
    with db() as c: return [dict(x) for x in c.execute('SELECT * FROM vacancies ORDER BY id DESC')]

class Vacancy(BaseModel):
    department:str
    title_uz:str=Field(min_length=2,max_length=150)
    title_zh:str=Field(min_length=2,max_length=150)
    description_uz:str=Field(default='',max_length=5000)
    description_zh:str=Field(default='',max_length=5000)
    active:bool=True

@app.post('/api/admin/vacancies')
def add_vac(v:Vacancy,uid:int=__import__('fastapi').Depends(admin)):
    if v.department not in ('Savdo','Ishlab chiqarish'): raise HTTPException(422,'Bo‘lim noto‘g‘ri')
    with db() as c:
        cur=c.execute('INSERT INTO vacancies (department,title_uz,title_zh,description_uz,description_zh,active) VALUES (?,?,?,?,?,?)',(v.department,v.title_uz,v.title_zh,v.description_uz,v.description_zh,int(v.active)))
        return {'id':cur.lastrowid}

@app.put('/api/admin/vacancies/{vid}')
def edit_vac(vid:int,v:Vacancy,uid:int=__import__('fastapi').Depends(admin)):
    if v.department not in ('Savdo','Ishlab chiqarish'): raise HTTPException(422,'Bo‘lim noto‘g‘ri')
    with db() as c:
        cur=c.execute('UPDATE vacancies SET department=?,title_uz=?,title_zh=?,description_uz=?,description_zh=?,active=? WHERE id=?',(v.department,v.title_uz,v.title_zh,v.description_uz,v.description_zh,int(v.active),vid))
        if not cur.rowcount: raise HTTPException(404,'Vakansiya topilmadi')
    return {'ok':True}

class Extra(BaseModel):
    question_uz:str=Field(min_length=2,max_length=500)
    question_zh:str=Field(min_length=2,max_length=500)

@app.post('/api/admin/vacancies/{vid}/questions')
def add_question(vid:int,q:Extra,uid:int=__import__('fastapi').Depends(admin)):
    with db() as c:
        if not c.execute('SELECT 1 FROM vacancies WHERE id=?',(vid,)).fetchone(): raise HTTPException(404,'Vakansiya yo‘q')
        cur=c.execute('INSERT INTO extra_questions (vacancy_id,question_uz,question_zh) VALUES (?,?,?)',(vid,q.question_uz,q.question_zh))
        return {'id':cur.lastrowid}

@app.delete('/api/admin/questions/{qid}')
def delete_question(qid:int,uid:int=__import__('fastapi').Depends(admin)):
    with db() as c: c.execute('DELETE FROM extra_questions WHERE id=?',(qid,))
    return {'ok':True}

def read_answers(s):
    try:
        value=json.loads(s)
        if isinstance(value,dict):return value
    except (ValueError,TypeError): pass
    # legacy bot stored repr(dict): safely parse, no eval
    import ast
    try:
        value=ast.literal_eval(s)
        return value if isinstance(value,dict) else {}
    except (ValueError,SyntaxError,TypeError): return {}

def pdf_bytes(app_id):
    with db() as c:
        row=c.execute('SELECT a.*,v.title_zh,v.title_uz,v.department FROM applications a LEFT JOIN vacancies v ON a.vacancy_id=v.id WHERE a.id=?',(app_id,)).fetchone()
        photo=c.execute('SELECT photo,mime FROM web_photos WHERE application_id=?',(app_id,)).fetchone()
    if not row: raise HTTPException(404,'Ariza topilmadi')
    ans=read_answers(row['answers'])
    pdfmetrics.registerFont(UnicodeCIDFont('STSong-Light'))
    buf=BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=(595,842),leftMargin=38,rightMargin=38,topMargin=34,bottomMargin=35)
    st=ParagraphStyle('zh',fontName='STSong-Light',fontSize=10,leading=17,spaceAfter=7,wordWrap='CJK')
    title=ParagraphStyle('titlezh',parent=st,fontSize=17,leading=23,spaceAfter=16)
    content=[Paragraph('BNBM Central Asia — 应聘登记表',title),Paragraph(f'申请编号：{app_id}　　岗位：{escape(row["title_zh"] or row["title_uz"] or "")}',st),Paragraph(f'申请日期：{escape(row["created_at"])}',st)]
    for key in FIELDS:
        value=str(ans.get(key,''))
        content.append(Paragraph(f'{escape(LABELS_ZH.get(key,key))}：{escape(TRANSLATE.get(value,value))}',st))
    for key,value in ans.items():
        if key not in FIELDS:
            label=LABELS_ZH.get(key,key)
            content.append(Paragraph(f'{escape(label)}：{escape(TRANSLATE.get(str(value),str(value)))}',st))
    if photo:
        try:
            from PIL import Image as PILImage
            pic=PILImage.open(BytesIO(photo['photo']))
            pic.thumbnail((220,220))
            im=BytesIO();pic.convert('RGB').save(im,format='JPEG');im.seek(0)
            content.extend([Spacer(1,12),Paragraph('个人照片：',st),Image(im,width=pic.width,height=pic.height)])
        except Exception: logging.exception('PDF photo failed')
    doc.build(content)
    return buf.getvalue()

async def telegram_notify(app_id,photo):
    if not TOKEN: return
    recipients={ADMIN_ID}|HR_IDS
    if not any(recipients): return
    try:
        data=pdf_bytes(app_id)
        async with httpx.AsyncClient(timeout=25) as client:
            for uid in recipients:
                if not uid: continue
                try:
                    response=await client.post(f'https://api.telegram.org/bot{TOKEN}/sendDocument',data={'chat_id':uid,'caption':f'BNBM HR — yangi anketa №{app_id}'},files={'document':(f'BNBM_anketa_{app_id}.pdf',data,'application/pdf')})
                    response.raise_for_status()
                    if not response.json().get('ok'): logging.error('Telegram send failed: %s',response.text[:200])
                except Exception: logging.exception('Telegram notification failed for %s',uid)
    except Exception: logging.exception('PDF creation failed')

@app.post('/api/applications')
async def apply(vacancy_id:int=Form(...),language:str=Form(...),answers:str=Form(...),photo:UploadFile=File(...),consent:str=Form(...),uid:int=__import__('fastapi').Depends(auth)):
    if consent!='yes': raise HTTPException(422,'Shaxsiy ma’lumotlarni qayta ishlashga rozilik zarur')
    if language not in ('uz','zh'): raise HTTPException(422,'Til noto‘g‘ri')
    if len(answers)>16000: raise HTTPException(413,'Anketa juda uzun')
    try: a=json.loads(answers)
    except ValueError: raise HTTPException(422,'Javoblar JSON emas')
    if not isinstance(a,dict) or any(not isinstance(k,str) or not isinstance(v,str) or len(v)>2000 for k,v in a.items()): raise HTTPException(422,'Javoblar noto‘g‘ri')
    if any(not a.get(k,'').strip() for k in FIELDS): raise HTTPException(422,'Majburiy savollar to‘ldirilmagan')
    try: datetime.strptime(a['birth_date'],'%d.%m.%Y')
    except ValueError: raise HTTPException(422,'Tug‘ilgan sana noto‘g‘ri')
    if not re.fullmatch(r'\+?[0-9 ()-]{9,20}',a['phone']): raise HTTPException(422,'Telefon noto‘g‘ri')
    with db() as c:
        vac=c.execute('SELECT * FROM vacancies WHERE id=? AND active=1',(vacancy_id,)).fetchone()
        if not vac: raise HTTPException(404,'Vakansiya yopilgan')
        extras=c.execute('SELECT id FROM extra_questions WHERE vacancy_id=?',(vacancy_id,)).fetchall()
    if 'kran' in vac['title_uz'].lower() and (not a.get('crane_type') or not a.get('certificate')): raise HTTPException(422,'Kran savollari to‘ldirilmagan')
    if any(not a.get(f'extra_{r["id"]}') for r in extras): raise HTTPException(422,'Qo‘shimcha savollar to‘ldirilmagan')
    raw=await photo.read(MAX_PHOTO+1)
    if len(raw)>MAX_PHOTO: raise HTTPException(413,'Rasm 5 MB dan oshmasin')
    try:
        from PIL import Image as PILImage
        im=PILImage.open(BytesIO(raw));im.verify()
        im=PILImage.open(BytesIO(raw))
        if im.width*im.height>16_000_000: raise ValueError('large image')
        im.thumbnail((1200,1200));out=BytesIO();im.convert('RGB').save(out,format='JPEG',quality=82);raw=out.getvalue()
    except Exception: raise HTTPException(422,'Yaroqli rasm tanlang')
    with db() as c:
        cur=c.execute('INSERT INTO applications (telegram_id,language,vacancy_id,answers,photo_file_id,created_at) VALUES (?,?,?,?,?,?)',(uid,language,vacancy_id,json.dumps(a,ensure_ascii=False),'',datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S UTC')))
        app_id=cur.lastrowid
        c.execute('INSERT INTO web_photos(application_id,photo,mime) VALUES (?,?,?)',(app_id,raw,'image/jpeg'))
    await telegram_notify(app_id,raw)
    return {'id':app_id,'ok':True}

@app.get('/api/admin/applications')
def list_apps(uid:int=__import__('fastapi').Depends(admin)):
    with db() as c:
        rows=c.execute('SELECT a.id,a.telegram_id,a.language,a.status,a.created_at,a.answers,v.title_uz,v.title_zh FROM applications a LEFT JOIN vacancies v ON v.id=a.vacancy_id ORDER BY a.id DESC LIMIT 300').fetchall()
    return [{**dict(r),'answers':read_answers(r['answers'])} for r in rows]

class Status(BaseModel):
    status:str

@app.patch('/api/admin/applications/{app_id}/status')
def update_status(app_id:int,s:Status,uid:int=__import__('fastapi').Depends(admin)):
    if s.status not in ('Yangi','Ko‘rib chiqilmoqda','Suhbat','Qabul qilindi','Rad etildi'): raise HTTPException(422,'Status noto‘g‘ri')
    with db() as c:
        cur=c.execute('UPDATE applications SET status=? WHERE id=?',(s.status,app_id))
        if not cur.rowcount: raise HTTPException(404,'Ariza yo‘q')
    return {'ok':True}

@app.get('/api/admin/applications/{app_id}/pdf')
def get_pdf(app_id:int,uid:int=__import__('fastapi').Depends(admin)):
    return StreamingResponse(BytesIO(pdf_bytes(app_id)),media_type='application/pdf',headers={'Content-Disposition':f'attachment; filename="BNBM_anketa_{app_id}.pdf"'})

@app.get('/api/admin/export')
def export(uid:int=__import__('fastapi').Depends(admin)):
    with db() as c:
        rows=c.execute('SELECT a.*,v.title_uz FROM applications a LEFT JOIN vacancies v ON a.vacancy_id=v.id ORDER BY a.id DESC').fetchall()
    wb=Workbook();ws=wb.active;ws.title='Arizalar'
    ws.append(['ID','Telegram ID','Til','Lavozim','Holat','Sana',*FIELDS])
    for r in rows:
        a=read_answers(r['answers'])
        # avoid spreadsheet formula injection from applicant-provided fields
        def safe(x):
            s=str(x or '')
            return "'"+s if s.startswith(('=','+','-','@','\t','\r')) else s
        ws.append([r['id'],r['telegram_id'],r['language'],safe(r['title_uz']),r['status'],r['created_at'],*[safe(a.get(k,'')) for k in FIELDS]])
    b=BytesIO();wb.save(b);b.seek(0)
    return StreamingResponse(b,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':'attachment; filename="BNBM_arizalar.xlsx"'})

@app.get('/')
def home():return FileResponse(ROOT/'static'/'index.html')
app.mount('/static',StaticFiles(directory=ROOT/'static'),name='static')
