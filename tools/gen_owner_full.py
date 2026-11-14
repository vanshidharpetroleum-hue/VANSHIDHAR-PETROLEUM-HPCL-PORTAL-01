import openpyxl,datetime,re,html
R=openpyxl.load_workbook('/tmp/rc/Vanshidhar Register_Aug-2026 to _Complete.xlsx',data_only=True)
DS=openpyxl.load_workbook('/tmp/rc/Vanshidhar Daily_Sales_Report_Aug2026_to.xlsx',data_only=True)
def fd(v):
    if isinstance(v,datetime.datetime): return v.strftime('%d-%m-%Y')
    return '' if v is None else str(v)
def inr(v,dec=0):
    if v is None or v=='': return ''
    if isinstance(v,str):
        try: v=float(v.replace(',',''))
        except: return html.escape(v)
    neg=round(v,2 if dec else 0)<0; v=abs(v); s=('%.'+str(dec)+'f')%v; i,_,d=s.partition('.')
    if len(i)>3:
        h,t=i[:-3],i[-3:]; h=re.sub(r'(\d)(?=(\d\d)+$)',r'\1,',h); i=h+','+t
    return ('-' if neg else '')+i+('.'+d if d else '')
secs=[]
def table(title,note,heads,rows,kinds,tot=None):
    h='<h2 id="s%d">%d. %s</h2>'%(len(secs)+1,len(secs)+1,title)
    if note: h+='<div class=note>%s</div>'%note
    h+='<table><tr>'+''.join('<th class=%s>%s</th>'%('l' if k=='t' else 'r',x) for x,k in zip(heads,kinds))+'</tr>'
    for r in rows:
        h+='<tr>'+''.join('<td class=%s>%s</td>'%('l' if k=='t' else 'r', html.escape(fd(v)) if k=='t' else inr(v,2 if k=='n2' else 0)) for v,k in zip(r,kinds))+'</tr>'
    if tot: h+='<tr class=t>'+''.join('<td class=%s>%s</td>'%('l' if k=='t' else 'r', html.escape(str(v)) if isinstance(v,str) else inr(v,2 if k=='n2' else 0)) for v,k in zip(tot,kinds))+'</tr>'
    secs.append((title,h+'</table>'))
def rows_of(s,start,ncol,key=0):
    ws=R[s]; out=[]
    for r in range(start,ws.max_row+1):
        v=[c.value for c in ws[r][:ncol]]
        if isinstance(v[key],datetime.datetime) and any(x not in (None,0,'') for x in v[1:]): out.append(v)
    return out
num=lambda x: x if isinstance(x,(int,float)) else 0
# 1 DSR collection date-wise
rows=[]
for name in DS.sheetnames:
    m=re.match(r'(\d\d)-(\d\d)-(\d\d)$',name)
    if not m: continue
    ws=DS[name]; dt=ws['B4'].value
    if not isinstance(dt,datetime.datetime) or not isinstance(ws['D8'].value,(int,float)): continue
    if not ws['B33'].value: continue
    rows.append([dt,ws['F14'].value,ws['F15'].value,ws['H16'].value,ws['B26'].value,ws['B27'].value,ws['B28'].value,ws['B29'].value,ws['B30'].value,ws['B31'].value,ws['B32'].value,ws['B38'].value])
rows.sort(key=lambda r:r[0])
tot=['KUL']+[sum(num(r[i]) for r in rows) for i in range(1,12)]
table('Roz ki bikri aur collection (DSR ke hisab se)','Daily Sales file ke har din ki sheet se. Short plus = paisa kam, minus = zyada.',
 ['Date','MS L','HSD L','Bikri ₹','Cash','UPI','Owner','Party','Staff','Kharcha','Udhar','Short'],rows,['t','n2','n2','n','n','n','n','n','n','n','n','n'],tot)
# 2 Daily Report register
rows=[r for r in rows_of('Daily Report',4,9) if num(r[2]) or num(r[4])]; rows=[[r[0],r[2],r[3],r[4],r[5],r[6],r[7],r[8]] for r in rows]
table('Roz ki bikri (Register: Daily Report, 27 Aug se)','',['Date','MS L','MS ₹','HSD L','HSD ₹','Kul L','Kul ₹','Commission ₹'],rows,['t','n2','n','n2','n','n2','n','n'],['KUL']+[sum(num(r[i]) for r in rows) for i in range(1,8)])
# 3 stock
rows=rows_of('Fuel Stock Ledger',3,11)
table('Fuel Stock (date-wise)','',['Date','MS Opening','MS Kharid','MS Bikri','MS Closing','MS Value ₹','HSD Opening','HSD Kharid','HSD Bikri','HSD Closing','HSD Value ₹'],rows,['t']+['n2']*10)
# 4 oil purchase
rows=rows_of('Oil Purchage',3,6); rows=[[r[0],r[1],r[2],r[3],r[4],r[5]] for r in rows]
table('Tel kharid (HPCL Oil Purchase)','',['Date','Product','Invoice','Litre','Rate','Amount ₹'],rows,['t','t','t','n','n2','n2'],['KUL','','',sum(num(r[3]) for r in rows),'',sum(num(r[5]) for r in rows)])
# 5 owner
rows=rows_of('Owner Personal Book',3,6); rows=[[r[0],r[2],r[3],r[4],r[5]] for r in rows]
table('Owner ka len-den (Ashwini Kumar Singh)','Credit = owner ne pump me diya, Debit = pump se owner ke liye gaya.',['Date','Vivaran','Credit ₹','Debit ₹','Balance ₹'],rows,['t','t','n','n','n'],['KUL','',sum(num(r[2]) for r in rows),sum(num(r[3]) for r in rows),''])
# 6 party
rows=rows_of('Party Credit Accounts',3,6); rows=[[r[0],r[1],r[2],r[3],r[4],r[5]] for r in rows]
table('Party ka len-den (Party Credit Accounts)','Credit = party se paisa liya (pump ko dena), Debit = party ko wapas diya. Balance us party ka chalta balance.',['Date','Party','Vivaran','Credit ₹','Debit ₹','Balance ₹'],rows,['t','t','t','n','n','n'])
# 7 udhar
rows=rows_of('Credit Sales',4,6)
table('Credit Sale (Udhar) date-wise','Baaki = udhar diya, Paid = wapas mila. Chalta kul us party ka.',['Date','Party','Vivaran','Udhar ₹','Jama ₹','Party baaki ₹'],rows,['t','t','t','n','n','n'],['KUL','','',sum(num(r[3]) for r in rows),sum(num(r[4]) for r in rows),sum(num(r[3]) for r in rows)-sum(num(r[4]) for r in rows)])
# 8 staff
rows=rows_of('Staff Wages Ledger',3,5); rows=[[r[0],r[1],r[2],r[3],r[4]] for r in rows]
table('Staff payment / advance','',['Date','Staff','Vivaran','Rakam ₹','Mode'],rows,['t','t','t','n','t'],['KUL','','',sum(num(r[3]) for r in rows),''])
# 9 pump costs
rows=rows_of('Pump Miscellaneous Costs',3,5)
table('Pump kharcha','',['Date','Kis par','Vivaran','Rakam ₹','Mode'],rows,['t','t','t','n','t'],['KUL','','',sum(num(r[3]) for r in rows),''])
# 10 interest
ws=R['Debit Interest Book']; rows=[]
for r in range(3,ws.max_row+1):
    v=[c.value for c in ws[r][:5]]
    if v[0] and v[3] not in (None,''): rows.append([v[0],v[1],v[2],v[3],v[4]])
table('Debit Interest / bank charge','',['Date','Account','Vivaran','Rakam ₹','Prakar'],rows,['t','t','t','n2','t'])
# 11 SBI CC
rows=rows_of('SBI CC Bank Ledger',3,6); rows=[[r[0],r[2],r[3],r[4],r[5]] for r in rows]
table('SBI CC Bank Ledger','',['Date','Vivaran','Jama ₹','Nikasi ₹','Balance ₹'],rows,['t','t','n2','n2','n2'],['KUL','',sum(num(r[2]) for r in rows),sum(num(r[3]) for r in rows),rows[-1][4]])
# 12 HPCL
rows=rows_of('HPCL ',3,6); rows=[[r[0],r[1],r[2],r[3],r[4],r[5]] for r in rows]
table('HPCL account','Credit = HPCL ko paisa bheja, Debit = tel ka bill.',['Date','Vivaran','Ref','Credit ₹','Debit ₹','Balance ₹'],rows,['t','t','t','n2','n2','n2'])
# 13 mobil
rows=rows_of('Mobil Lubes Log',3,7)
table('Mobil / Lube','',['Date','Item','Opening','Kharid','Bikri','Closing','Aamdani ₹'],rows,['t','t','n','n','n','n','n'],['KUL','','','','','',sum(num(r[6]) for r in rows)])
toc='<div class=toc><b>Isme kya hai:</b><ol>'+''.join('<li><a href="#s%d">%s</a></li>'%(i+1,t) for i,(t,_) in enumerate(secs))+'</ol></div>'
css=open('/home/claude/vanshidhar-petroleum-hpcl-portal-01/tools/report_style.css').read() if False else ''
H='''<!DOCTYPE html><html><head><meta charset=utf-8><style>
body{font-family:Arial,sans-serif;color:#1d2433;margin:0}.page{padding:18px 22px}
.hd{background:#0d4f8b;color:#fff;border-radius:10px;padding:14px 18px}.hd h1{margin:0;font-size:22px}.hd .s{font-size:12px;opacity:.9;margin-top:4px}
h2{color:#0d4f8b;font-size:15px;border-bottom:2px solid #dfe6ef;padding-bottom:4px;margin:20px 0 6px;break-after:avoid}
table{width:100%;border-collapse:collapse;font-size:9.5px}th{background:#eef3f9}th,td{border:1px solid #d5dde8;padding:3px 4px}
tr{break-inside:avoid}.l{text-align:left}.r{text-align:right;white-space:nowrap}tr.t td{font-weight:bold;background:#f4f7fb}
.note{font-size:10px;color:#556;margin:2px 0 6px}.toc{font-size:12px;margin-top:12px}.toc a{color:#0d4f8b}
</style></head><body><div class=page><div class=hd><h1>OWNER REPORT - Poora date-wise hisaab</h1><div class=s>Vanshidhar Petroleum, HPCL Pansa · Shri Ashwini Kumar Singh ji ke liye · 1 Aug 2026 se 14-11-2026 tak · Register aur Daily Sales file se · Taiyar: Dharmendra Kumar Singh (Manager)</div></div>'''+toc+''.join(h for _,h in secs)+'</div></body></html>'
open('/home/claude/full/full.html','w').write(H); print([ (t,h.count('<tr>')) for t,h in secs])
