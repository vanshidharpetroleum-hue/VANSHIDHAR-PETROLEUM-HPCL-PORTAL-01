"""Ek din ki DSR report (PNG + PDF). Chalana: python tools/gen_dsr.py 12-11-26"""
import sys,os,re,subprocess,tempfile,datetime,openpyxl
ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__))); SN=sys.argv[1]
tmp=tempfile.mkdtemp(); prof=os.path.join(tmp,'p','user'); os.makedirs(prof)
open(os.path.join(prof,'registrymodifications.xcu'),'w').write('<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item></oor:items>')
src=os.path.join(ROOT,'Vanshidhar Daily_Sales_Report_Aug2026_to.xlsx')
subprocess.run(['soffice','-env:UserInstallation=file://'+os.path.join(tmp,'p'),'--headless','--convert-to','xlsx','--outdir',tmp,src],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
ws=openpyxl.load_workbook(os.path.join(tmp,os.path.basename(src)),data_only=True)[SN]
v=lambda c: ws[c].value
def n(x,d=2):
    if x in (None,''): return '-'
    neg=x<0; s=f"{abs(x):,.{d}f}"; i,_,f=s.partition('.'); i=i.replace(',','')
    if len(i)>3: h,t=i[:-3],i[-3:]; h=re.sub(r'(\d)(?=(\d\d)+$)',r'\1,',h); i=h+','+t
    return ('-' if neg else '')+i+('.'+f if f else '')
dt=v('B4'); DS=dt.strftime('%d-%m-%Y'); day=['Somvar','Mangalvar','Budhvar','Guruvar','Shukravar','Shanivar','Ravivar'][dt.weekday()]
noz=''.join(f"<tr><td>{v('A%d'%r)}</td><td>{v('B%d'%r)}</td><td>{n(v('C%d'%r))}</td><td>{n(v('D%d'%r))}</td><td>{n(v('E%d'%r))}</td><td>{n(v('F%d'%r))}</td><td>{v('G%d'%r)}</td><td>{n(v('H%d'%r))}</td></tr>" for r in range(8,14))
def stock(r,lab):
    dip=v('F%d'%r); var=(dip-v('E%d'%r)) if isinstance(dip,(int,float)) and dip else None
    return f"<tr><td>{lab}</td><td>{n(v('B%d'%r))}</td><td>{n(v('C%d'%r) or 0)}</td><td>{n(v('D%d'%r))}</td><td>{n(v('E%d'%r))}</td><td>{n(dip) if dip else 'Pending'}</td><td class={'grn' if (var or 0)>=0 else 'red'}>{(('+' if var>=0 else '')+n(var)) if var is not None else '-'}</td></tr>"
coll=''.join(f"<tr><td>{v('A%d'%r)}</td><td>{n(v('B%d'%r))}</td><td style='text-align:left'>{v('C%d'%r) or ''}</td></tr>" for r in range(26,33) if v('B%d'%r))
sh=v('B38') or 0
CSS=open(os.path.join(ROOT,'tools','report_style.css')).read()
html=f"""<!doctype html><html><head><meta charset=utf-8><style>{CSS}</style></head><body><div class=page>
<div class=hd><h1>VANSHIDHAR PETROLEUM - HPCL PANSA</h1><div class=s>DAILY SALES REPORT (DSR) · Date: {DS} ({day}) · Manager: Dharmendra Kumar Singh</div></div>
<div class=cards><div class=c><div class=l>KUL BIKRI</div><div class=v>₹{n(v('B37'))}</div></div><div class=c><div class=l>MS / HSD (Litre)</div><div class=v>{n(v('F14'))} / {n(v('F15'))}</div></div><div class=c><div class=l>SHORT / (EXCESS)</div><div class="v {'red' if sh>0 else 'grn'}">₹{n(sh)}{' (Excess)' if sh<0 else ''}</div></div></div>
<h2>A. Nozzle / Meter Reading</h2><table><tr><th>Nozzle</th><th>Product</th><th>Opening</th><th>Closing</th><th>Testing (L)</th><th>Net sale (L)</th><th>Rate</th><th>Amount ₹</th></tr>{noz}
<tr class=t><td colspan=5>Total MS (Petrol)</td><td>{n(v('F14'))}</td><td></td><td>{n(v('H14'))}</td></tr><tr class=t><td colspan=5>Total HSD (Diesel)</td><td>{n(v('F15'))}</td><td></td><td>{n(v('H15'))}</td></tr><tr class=t><td colspan=5>TOTAL FUEL SALE</td><td>{n(v('F16'))}</td><td></td><td>{n(v('H16'))}</td></tr></table>
<h2>B. Stock (Tank)</h2><table><tr><th>Product</th><th>Opening (L)</th><th>Tanker (L)</th><th>Sale (L)</th><th>Book closing (L)</th><th>Asli dip (L)</th><th>Variation</th></tr>{stock(20,'MS (Petrol)')}{stock(21,'HSD (Diesel)')}</table>
<h2>C. Collection</h2><table><tr><th>Mode</th><th>Amount ₹</th><th style="text-align:left">Remarks</th></tr>{coll}
<tr class=t><td>TOTAL COLLECTED</td><td>{n(v('B33'))}</td><td></td></tr><tr class=t><td>Fuel sale</td><td>{n(v('B35'))}</td><td></td></tr>
{f"<tr class=t><td>Lube / Mobil</td><td>{n(v('B36'))}</td><td style='text-align:left'>{v('C36') or ''}</td></tr>" if v('B36') else ''}
<tr class=t><td>TOTAL EXPECTED</td><td>{n(v('B37'))}</td><td></td></tr><tr class=t><td>SHORT / (EXCESS)</td><td class={'red' if sh>0 else 'grn'}>{n(sh)}</td><td style="text-align:left">Plus = paisa kam aaya, Minus = zyada aaya</td></tr></table>
<div class=ft><span>Taiyar: {datetime.date.today().strftime('%d-%m-%Y')}</span><span>Manager hastakshar: ____________</span></div></div></body></html>"""
out=os.path.join(ROOT,'Reports','DSR_'+DS)
open(out+'.html','w').write(html)
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':794,'height':1000},device_scale_factor=2); pg.goto('file://'+out+'.html')
    pg.screenshot(path=out+'.png',full_page=True); pg.pdf(path=out+'.pdf',format='A4',print_background=True,margin={'top':'0','bottom':'0','left':'0','right':'0'}); b.close()
os.remove(out+'.html'); print(out)
