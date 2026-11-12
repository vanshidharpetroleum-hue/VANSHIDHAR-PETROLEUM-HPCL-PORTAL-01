"""index.html (Portal) ka DATA Register file se naya banata hai. Chalana: python tools/build_portal.py [recalculated_register.xlsx]"""
import json,re,os,sys,datetime,subprocess,tempfile,openpyxl
ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REGN='Vanshidhar Register_Aug-2026 to _Complete.xlsx'
if len(sys.argv)>1: path=sys.argv[1]
else:
    tmp=tempfile.mkdtemp(); prof=os.path.join(tmp,'p','user'); os.makedirs(prof)
    open(os.path.join(prof,'registrymodifications.xcu'),'w').write('<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>')
    subprocess.run(['soffice','-env:UserInstallation=file://'+os.path.join(tmp,'p'),'--headless','--convert-to','xlsx','--outdir',tmp,os.path.join(ROOT,REGN)],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    path=os.path.join(tmp,REGN)
wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
def cv(v):
    if v is None: return ''
    if isinstance(v,datetime.datetime): return v.strftime('%Y-%m-%d')
    if isinstance(v,float): return round(v,2)
    return v
def grab(sn):
    rows=[[cv(c) for c in r] for r in wb[sn].iter_rows(values_only=True)]
    rows=[r for r in rows if any(x!='' for x in r)]
    w=max((max([i+1 for i,x in enumerate(r) if x!=''] or [0]) for r in rows),default=0)
    return [ (r+['']*w)[:w] for r in rows]
SHEETS=[('Deshbord','Deshbord'),('Monthly Summery','Monthly Summary'),('Yearly Report','Yearly'),('Daily Report','Daily Report'),('Fuel Stock Ledger','Fuel Stock Ledger'),
 ('Nozzle Sales Logs','Nozzle Sales Logs'),('Oil Purchage','Oil Purchage'),('SBI CC Bank Ledger','SBI CC Bank Ledger'),('Pump Miscellaneous Costs','Pump Miscellaneous Costs'),
 ('Staff Wages Ledger','Staff Wages Ledger'),('Debit Interest Book','Debit Interest Book'),('Owner Personal Book','Owner Personal Book'),('Owner & Party Monthly','Owner & Party Monthly'),
 ('Party Credit Accounts','Party Credit Accounts'),('Credit Sales','Credit Sales (Udhar)'),('HPCL ','HPCL'),('Mobil Lubes Log','Mobil Lubes Log'),('P&L Statement','P&L Statement')]
sheets={label:grab(sn) for sn,label in SHEETS if sn in wb.sheetnames}
metrics={}
for r in wb['Deshbord'].iter_rows(min_row=3,values_only=True):
    if r and isinstance(r[0],str) and isinstance(r[1],(int,float)): metrics[r[0].strip()]=round(r[1],2)
# last data date from Fuel Stock Ledger
last=None
for r in wb['Fuel Stock Ledger'].iter_rows(values_only=True):
    if r and isinstance(r[0],datetime.datetime) and isinstance(r[3],(int,float)) and r[3]>0 and (last is None or r[0]>last): last=r[0]
DATA={'metrics':metrics,'sheets':sheets,'updated':(last or datetime.datetime.now()).strftime('%d-%m-%Y')}
f=os.path.join(ROOT,'index.html'); s=open(f,encoding='utf-8').read()
s=re.sub(r'const DATA=\{.*?\};\n',lambda m:'const DATA='+json.dumps(DATA,ensure_ascii=False)+';\n',s,count=1,flags=re.S)
s=re.sub(r'<div class="sub">.*?</div></header>','<div class="sub">Management Portal • Data: 24 Aug se '+DATA['updated']+' tak • <a href="dashboard.html" style="color:#ffd34a">📊 Dashboard</a> • <a href="https://github.com/vanshidharpetroleum-hue/VANSHIDHAR-PETROLEUM-HPCL-PORTAL-01/tree/main/Reports" style="color:#ffd34a">📄 Reports</a></div></header>',s,count=1,flags=re.S)
open(f,'w',encoding='utf-8').write(s); print('portal', DATA['updated'], len(sheets), list(metrics)[:3])
