import openpyxl,datetime,re,collections,html
w=openpyxl.load_workbook('/tmp/rc/Vanshidhar Register_Aug-2026 to _Complete.xlsx',data_only=True)
def inr(v):
    neg=v<0; s='%.0f'%abs(v)
    if len(s)>3: s=re.sub(r'(\d)(?=(\d\d)+$)',r'\1,',s[:-3])+','+s[-3:]
    return ('-' if neg else '')+s
p=w['Party Credit Accounts']; bal={}; last={}; nm={}
for r in range(3,120):
    b=p.cell(r,2).value; f=p.cell(r,6).value; a=p.cell(r,1).value
    if b and isinstance(f,(int,float)):
        k=re.sub(r'\s+',' ',b.strip().upper()); bal[k]=f; nm[k]=b.strip().title(); last[k]=a
cs=w['Credit Sales']; ud=collections.OrderedDict(); ul={}; un={}
for r in range(4,62):
    b=cs.cell(r,2).value; a=cs.cell(r,1).value
    if not b or not isinstance(a,datetime.datetime): continue
    k=re.sub(r'\s+',' ',b.strip().upper()); d=cs.cell(r,4).value or 0; e=cs.cell(r,5).value or 0
    ud[k]=ud.get(k,0)+(d if isinstance(d,(int,float)) else 0)-(e if isinstance(e,(int,float)) else 0); ul[k]=a; un[k]=b.strip()
dena=sorted([(nm[k],v,last[k]) for k,v in bal.items() if v>0.5],key=lambda x:-x[1])
lena_party=sorted([(nm[k],-v,last[k]) for k,v in bal.items() if v<-0.5],key=lambda x:-x[1])
lena=sorted([(un[k],v,ul[k]) for k,v in ud.items() if v>0.5],key=lambda x:-x[1])
extra=[(un[k],v) for k,v in ud.items() if v<-0.5]
fd=lambda a:a.strftime('%d-%m-%Y') if isinstance(a,datetime.datetime) else ''
interest={'CHOTU SINGH','GUDU SINGH','MOON SINGH'}
def tbl(rows,head,tot,note=lambda n:''):
    h='<table><tr><th>#</th><th class=l>%s</th><th class=r>Baaki ₹</th><th>Aakhri entry</th><th class=l>Note</th></tr>'%head
    for i,(n,v,a) in enumerate(rows,1): h+='<tr><td>%d</td><td class=l>%s</td><td class=r>%s</td><td>%s</td><td class=l>%s</td></tr>'%(i,html.escape(n),inr(v),fd(a),note(n))
    return h+'<tr class=t><td></td><td class=l>KUL</td><td class=r>%s</td><td></td><td></td></tr></table>'%inr(tot)
D=sum(v for _,v,_ in dena); L=sum(v for _,v,_ in lena); LP=sum(v for _,v,_ in lena_party)
newp={'VIVEK VERMA','JAIKUSH SINGH','GOLU SINGH','RAVISH SINGH','SURAJ MEHTA'}
dnote=lambda n:'3% byaj' if n.upper() in interest else ('Nov me liya' if n.upper() in newp else ('Lena list me bhi hai - check' if 'DOOARS' in n.upper() else ''))
def lnote(n):
    if any(n==x for x,_,_ in lena_party): return 'Party khata (zyada diya gaya)'
    if 'TUTU' in n.upper() or 'ROUSHAN' in n.upper(): return 'Purana udhar (2021)'
    if 'DOOARS' in n.upper(): return 'Dena list me bhi hai - check'
    return 'Chalu udhar'
H='''<!DOCTYPE html><html><head><meta charset=utf-8><style>
body{font-family:Arial,sans-serif;color:#1d2433;margin:0;font-size:12px}.page{padding:16px 24px}
.hd{background:#0d4f8b;color:#fff;border-radius:10px;padding:14px 18px}.hd h1{margin:0;font-size:21px}.hd .s{font-size:12px;opacity:.9;margin-top:4px}
h2{font-size:15px;border-bottom:2px solid #dfe6ef;padding-bottom:4px;margin:18px 0 8px;break-after:avoid}
table{width:100%%;border-collapse:collapse;font-size:11.5px}th{background:#eef3f9}th,td{border:1px solid #d5dde8;padding:4px 6px;text-align:center}
.l{text-align:left}.r{text-align:right;white-space:nowrap}tr.t td{font-weight:bold;background:#f4f7fb}tr{break-inside:avoid}
.kpis{display:flex;gap:8px;margin-top:12px}.k{flex:1;border:1px solid #d5dde8;border-radius:8px;padding:8px 10px}.k b{display:block;font-size:18px;margin-top:2px}
.red{color:#b42318}.grn{color:#067647}.small{font-size:10px;color:#556;margin-top:4px}
</style></head><body><div class=page>
<div class=hd><h1>PARTY KA DENA aur LENA - LIST</h1><div class=s>Vanshidhar Petroleum, HPCL Pansa · 15-11-2026 tak · Register (Party Credit Accounts aur Credit Sales) se</div></div>
<div class=kpis><div class=k>Pump ko DENA hai (karja)<b class=red>₹%s</b></div><div class=k>Pump ko LENA hai (udhar)<b class=grn>₹%s</b></div><div class=k>Fark (dena − lena)<b class=red>₹%s</b></div></div>
<h2 class=red style="color:#b42318">A. Pump ko DENA hai – %d party</h2>%s
<div class=small>Inse pump ne paisa liya hai (karja). "3%% byaj" wali party ko har mahina byaj dena padta hai.</div>
<h2 style="color:#067647">B. Pump ko LENA hai (Credit Sale / Udhar) – %d party</h2>%s
'''%(inr(D),inr(L+LP),inr(D-L-LP),len(dena),tbl(dena,'Party (kise dena hai)',D,dnote),len(lena)+len(lena_party),tbl(lena+lena_party,'Party (kisse lena hai)',L+LP,lnote))
if extra: H+='<div class=small>Zyada jama (round-off): '+', '.join('%s ₹%s'%(n,inr(-v)) for n,v in extra)+'</div>'
H+='<div class=small>Lena me: chalu udhar ₹91,766 + purana udhar 2021 (Tutu Singh 1,50,000, Roushan Singh 20,000) + party khate me zyada diya (Ravi Singh 14,443, J P Const 131). Dooars Construction dono list me hai, ek baar milan karna hai.</div></div></body></html>'
open('lists.html','w').write(H); print(len(dena),D,len(lena),L,lena_party,extra)

CSS=H[:H.index('<div class=hd>')]
def page(title,sub,kpi,body):
    return CSS+'<div class=hd><h1>%s</h1><div class=s>%s</div></div><div class=kpis>%s</div>%s</div></body></html>'%(title,sub,kpi,body)
sub='Vanshidhar Petroleum, HPCL Pansa · 15-11-2026 tak · Register se'
dh=page('PUMP KO DENA HAI - PARTY LIST',sub,'<div class=k>Kul dena (%d party)<b class=red>₹%s</b></div><div class=k>3%% byaj wale (Chotu, Gudu, Moon)<b class=red>₹%s</b></div>'%(len(dena),inr(D),inr(sum(v for n,v,_ in dena if n.upper() in interest))),
  tbl(dena,'Party (kise dena hai)',D,dnote)+'<div class=small>Inse pump ne paisa liya hai (karja). 3% byaj wali party ko har mahina byaj dena padta hai. Dooars Construction lena list me bhi hai, milan baaki.</div>')
chalu=sum(v for n,v,_ in lena if lnote(n) in ('Chalu udhar','Dena list me bhi hai - check'))
lh=page('PUMP KO LENA HAI - UDHAR LIST',sub,'<div class=k>Kul lena (%d party)<b class=grn>₹%s</b></div><div class=k>Chalu udhar<b class=grn>₹%s</b></div><div class=k>Purana udhar (2021)<b>₹1,70,000</b></div>'%(len(lena)+len(lena_party),inr(L+LP),inr(chalu)),
  tbl(lena+lena_party,'Party (kisse lena hai)',L+LP,lnote)+'<div class=small>Chalu udhar = Nov-2026 tak ka credit sale. Purana udhar 2021 ka hai. Dooars Construction dena list me bhi hai, milan baaki.</div>')
open('dena.html','w').write(dh); open('lena.html','w').write(lh)
