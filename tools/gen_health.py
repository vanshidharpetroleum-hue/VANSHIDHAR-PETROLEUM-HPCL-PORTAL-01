"""Pump Health Report (har hafte). Chalane ka tarika: python tools/gen_health.py
dashboard.html (build_dashboard.py se) aur Register file se data leta hai, Reports/ me PNG + PDF banata hai."""
import json,re,os,sys,datetime,subprocess,tempfile,shutil,openpyxl
ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
s=open(os.path.join(ROOT,'dashboard.html'),encoding='utf-8').read()
D=json.loads(re.search(r'const D=(.*?);\nconst inr',s,re.S).group(1))
def recalc(name):
    tmp=tempfile.mkdtemp(); prof=os.path.join(tmp,'prof','user'); os.makedirs(prof)
    open(os.path.join(prof,'registrymodifications.xcu'),'w').write('<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>')
    subprocess.run(['soffice','-env:UserInstallation=file://'+os.path.join(tmp,'prof'),'--headless','--convert-to','xlsx','--outdir',os.path.join(tmp,'out'),os.path.join(ROOT,name)],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    return openpyxl.load_workbook(os.path.join(tmp,'out',os.path.basename(name)),read_only=True,data_only=True)
def ind(v):
    neg=v<0; i=str(int(round(abs(v))))
    if len(i)>3: h,t=i[:-3],i[-3:]; h=re.sub(r'(\d)(?=(\d\d)+$)',r'\1,',h); i=h+','+t
    return ('-' if neg else '')+i
R=lambda v:('-₹' if v<0 else '₹')+ind(abs(v))
daily=D['daily']; last=datetime.date.fromisoformat(daily[-1]['d'])
w0=last-datetime.timedelta(days=6); p0=w0-datetime.timedelta(days=7); p1=w0-datetime.timedelta(days=1)
def rng(a,b): return [x for x in daily if a<=datetime.date.fromisoformat(x['d'])<=b]
W,P=rng(w0,last),rng(p0,p1)
def agg(L):
    n=max(len(L),1); ms=sum(x['ms'] for x in L); hsd=sum(x['hsd'] for x in L)
    return dict(days=len(L),ms=ms,hsd=hsd,lit=ms+hsd,sales=sum(x['total'] for x in L),avg=sum(x['total'] for x in L)/n,avgl=(ms+hsd)/n,
        comm=ms*3.8+hsd*2.4,short=sum(x['short'] for x in L),credit=sum(x['credit'] for x in L),kharcha=sum(x['kharcha'] for x in L),
        upi=sum(x['upi'] for x in L),cash=sum(x['cash'] for x in L))
a,b=agg(W),agg(P)
def ch(k):
    if not b[k] or k=='short': return ''
    d=(a[k]-b[k])/abs(b[k])*100; good=(d>=0) if k not in ('credit','kharcha') else (d<=0); c='grn' if good else 'red'
    return f"<span class={c}>{'▲' if d>=0 else '▼'} {abs(d):.0f}%</span>"
# Register month
wb=recalc('Vanshidhar Register_Aug-2026 to _Complete.xlsx')
ms_=wb['Monthly Summery']; rows=list(ms_.iter_rows(values_only=True))
hdr=rows[1]; mi=[i for i,c in enumerate(hdr) if isinstance(c,datetime.datetime) and c.year==last.year and c.month==last.month]
def rv(name,col):
    for r in rows:
        if isinstance(r[0],str) and r[0].lower().startswith(name.lower()): return r[col] or 0
    return 0
c=mi[0] if mi else None; tot=hdr.index('Total')
M={k:rv(n,c) for k,n in [('sales','Total Sales'),('gross','Gross Profit'),('exp','Total General'),('staff','Total Staff'),('int','Total Debit Interest'),('op','Operating Profit'),('own','Owner Debit'),('party','Other Party Debit')]} if c is not None else {}
T={k:rv(n,tot) for k,n in [('sales','Total Sales'),('gross','Gross Profit'),('exp','Total General'),('staff','Total Staff'),('int','Total Debit Interest'),('op','Operating Profit'),('own','Owner Debit'),('party','Other Party Debit')]}
hpcl=rv('HPCL Credit Balance',tot)
udhar=None
for r in wb['Credit Sales'].iter_rows(values_only=True):
    if r and len(r)>3 and r[1]=='Udhar Balance': udhar=r[3]
# dips from Daily Summery
dw=recalc('Vanshidhar Daily_Sales_Report_Aug2026_to.xlsx')['Daily Summery']
nodip=[]
for r in dw.iter_rows(min_row=5,values_only=True):
    r=list(r)+[None]*20
    if isinstance(r[0],datetime.datetime) and w0<=r[0].date()<=last and (r[6] or 0)>0 and not ((r[15] or 0) and (r[16] or 0)): nodip.append(r[0].strftime('%d'))
# signals
def light(ok,warn=False): return '🟢' if ok else ('🟡' if warn else '🔴')
sig=[(light(a['avgl']>=b['avgl']*0.95 if b['days'] else True, a['avgl']>=b['avgl']*0.85 if b['days'] else True),'Bikri',f"Roz {ind(a['avgl'])} L ({R(a['avg'])}) {ch('avgl')}"),
     (light(abs(a['short'])<=500, abs(a['short'])<=2000),'Cash hisaab',f"Short/(Excess) hafte me {R(a['short'])}"),
     (light(len(nodip)==0, len(nodip)<=2),'Raat ka dip',('Sab din likha gaya' if not nodip else f"{len(nodip)} din dip nahi likha ({', '.join(nodip)} tarikh)")),
     (light(T['int']<=0, T['int']<50000),'HPCL byaj',f"Aug se ab tak {R(T['int'])}"),
     (light(T['op']>=0, T['op']>-50000),'Munafa (Operating)',f"Aug se ab tak {R(T['op'])}"),
     (light((udhar or 0)<75000,(udhar or 0)<150000),'Chalu udhar',R(udhar or 0))]
tips=[]
if T['int']>0: tips.append(f"HPCL ka payment time par kariye. Aug se ab tak {R(T['int'])} byaj laga hai, yeh seedha munafa khata hai.")
if nodip: tips.append("Roz raat ka asli dip DSR me likhiye. Bina dip ke stock ki kami pakad me nahi aati.")
if T['own']+T['party']>T['gross']: tips.append(f"Pump ke paise se owner ({R(T['own'])}) aur party ({R(T['party'])}) ka len-den kam kariye. Isi se HPCL ko udhar par maal lena padta hai.")
if a['ms']<a['hsd']: tips.append("MS (petrol) me commission zyada (₹3.8/L) hai. Bike/car walon ke liye offer ya service par dhyan dijiye.")
if abs(a['short'])>500: tips.append("Short/Excess ₹500 se zyada hai. Har shift ke ant me cash milan kariye.")
tips=tips[:4]
WS=w0.strftime('%d %b'); WE=last.strftime('%d %b %Y'); MN=last.strftime('%B %Y')
CSS="""*{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;color:#172033;background:#fff}.page{width:794px;padding:26px 32px}
.hd{background:#0b4f91;color:#fff;border-radius:10px;padding:14px 18px}.hd h1{margin:0;font-size:24px}.hd .s{font-size:13px;opacity:.9;margin-top:4px}
h2{font-size:15px;margin:16px 0 6px;color:#0b4f91;border-bottom:2px solid #dce3ec;padding-bottom:4px}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{border:1px solid #d5deea;padding:6px 8px;text-align:right}th{background:#eef3f9}td:first-child,th:first-child{text-align:left}
.sig td{font-size:14px}.red{color:#b42318;font-weight:bold}.grn{color:#11793a;font-weight:bold}
.cards{display:grid;grid-template-columns:repeat(4,1fr);gap:8px}.c{border:1px solid #dce3ec;border-radius:10px;padding:9px}.c .l{font-size:11px;color:#65748a;font-weight:bold}.c .v{font-size:18px;font-weight:800;margin-top:3px}
ol{margin:4px 0 0 18px;padding:0;font-size:13.5px;line-height:1.55}.note{font-size:11px;color:#65748a;margin-top:5px}.ft{margin-top:14px;font-size:11px;color:#65748a;display:flex;justify-content:space-between}"""
sigrows=''.join(f"<tr><td style='width:34px;text-align:center;font-size:18px'>{l}</td><td style='text-align:left;font-weight:bold'>{n}</td><td style='text-align:left'>{v}</td></tr>" for l,n,v in sig)
cmp=lambda lab,k,f: f"<tr><td>{lab}</td><td>{f(a[k])}</td><td>{f(b[k]) if b['days'] else '-'}</td><td>{ch(k)}</td></tr>"
L=lambda v: ind(v)+' L'
html=f"""<!doctype html><html><head><meta charset=utf-8><style>{CSS}</style></head><body><div class=page>
<div class=hd><h1>PUMP HEALTH REPORT</h1><div class=s>Vanshidhar Petroleum, HPCL Pansa · Hafta: {WS} se {WE} · Owner: Shri Ashwini Kumar Singh</div></div>
<h2>Ek nazar me</h2><table class=sig>{sigrows}</table>
<h2>Is hafte vs pichhla hafta</h2>
<table><tr><th>Cheez</th><th>Is hafte ({a['days']} din)</th><th>Pichhla hafta ({b['days']} din)</th><th>Badlav</th></tr>
{cmp('Kul bikri ₹','sales',R)}{cmp('Roz ki bikri ₹','avg',R)}{cmp('MS (Petrol)','ms',L)}{cmp('HSD (Diesel)','hsd',L)}{cmp('Roz ka litre','avgl',L)}
{cmp('Andaazit commission ₹','comm',R)}{cmp('Naya udhar diya ₹','credit',R)}{cmp('Staff advance + pump kharcha ₹','kharcha',R)}{cmp('Short / (Excess) ₹','short',R)}</table>
<div class=note>Commission ka andaaza: MS ₹3.8/L, HSD ₹2.4/L (Register ki Daily Report jaisa).</div>
<h2>{MN} ab tak (Register)</h2>
<div class=cards><div class=c><div class=l>BIKRI</div><div class=v>{R(M.get('sales',0))}</div></div><div class=c><div class=l>GROSS PROFIT</div><div class=v>{R(M.get('gross',0))}</div></div>
<div class=c><div class=l>KHARCHA+STAFF+BYAJ</div><div class=v>{R(M.get('exp',0)+M.get('staff',0)+M.get('int',0))}</div></div><div class=c><div class=l>OPERATING PROFIT</div><div class="v {'red' if M.get('op',0)<0 else 'grn'}">{R(M.get('op',0))}</div></div></div>
<h2>Aug-2026 se ab tak (Register)</h2>
<div class=cards><div class=c><div class=l>BIKRI</div><div class=v>{R(T['sales'])}</div></div><div class=c><div class=l>OPERATING PROFIT</div><div class="v {'red' if T['op']<0 else 'grn'}">{R(T['op'])}</div></div>
<div class=c><div class=l>HPCL BYAJ</div><div class="v red">{R(T['int'])}</div></div><div class=c><div class=l>CHALU UDHAR</div><div class=v>{R(udhar or 0)}</div></div>
<div class=c><div class=l>OWNER DEBIT</div><div class=v>{R(T['own'])}</div></div><div class=c><div class=l>OTHER PARTY DEBIT</div><div class=v>{R(T['party'])}</div></div>
<div class=c><div class=l>STAFF PAYMENT</div><div class=v>{R(T['staff'])}</div></div><div class=c><div class=l>HPCL CREDIT BALANCE</div><div class=v>{R(hpcl)}</div></div></div>
<h2>Is hafte ke sujhav</h2><ol>{''.join(f'<li>{x}</li>' for x in tips) or '<li>Sab theek chal raha hai. Isi tarah jaari rakhiye.</li>'}</ol>
<div class=ft><span>Taiyar: Dharmendra Kumar Singh (Manager) · {datetime.date.today().strftime('%d-%m-%Y')}</span><span>🟢 theek · 🟡 dhyan dein · 🔴 sudhar zaroori</span></div>
</div></body></html>"""
os.makedirs(os.path.join(ROOT,'Reports'),exist_ok=True)
out=os.path.join(ROOT,'Reports',f"Pump_Health_{last.strftime('%d-%m-%Y')}")
open(out+'.html','w',encoding='utf-8').write(html)
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    b_=p.chromium.launch(); pg=b_.new_page(viewport={'width':794,'height':1000},device_scale_factor=2)
    pg.goto('file://'+out+'.html'); pg.screenshot(path=out+'.png',full_page=True)
    pg.pdf(path=out+'.pdf',format='A4',print_background=True,margin={'top':'0','bottom':'0','left':'0','right':'0'}); b_.close()
os.remove(out+'.html'); print(out)
