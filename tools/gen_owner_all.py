import json,re,openpyxl,datetime,collections
REPO='/home/claude/vanshidhar-petroleum-hpcl-portal-01/'
REG='/tmp/lo8/out/r.xlsx'
s=open(REPO+'dashboard.html').read()
D=json.loads(re.search(r'const D=(.*?);\nconst inr',s,re.S).group(1))
def ind(v):
    if v is None: return '-'
    neg=v<0; i=str(int(round(abs(v))))
    if len(i)>3: h,t=i[:-3],i[-3:]; h=re.sub(r'(\d)(?=(\d\d)+$)',r'\1,',h); i=h+','+t
    return ('-' if neg else '')+i
R=lambda v:('-₹' if v<0 else '₹')+ind(abs(v))
def cr(v): return ('-' if v<0 else '')+('₹%.2f cr'%(abs(v)/1e7) if abs(v)>=1e7 else '₹%.2f lakh'%(abs(v)/1e5))
TODAY=D['daily'][-1]['d']; td=datetime.date.fromisoformat(TODAY); TDS=td.strftime('%d-%m-%Y'); TDN='%d %s'%(td.day,['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][td.month-1])
own={'2020-21':(6405000,7321999),'2021-22':(7694900,6209267),'2022-23':(4100000,4183050),'2023-24':(7305000,8783929),'2024-25':(4709529,2235421),'2025-26':(6068840,2373761),'2026-27':(1464000,3467872)}
yrs=[y for y in D['yearly'] if y['fy']!='TOTAL']
g=lambda n:next(r for r in D['reg']['rows'] if r['name'].lower().startswith(n.lower()))
rs,rp,rcg,rgp=g('Total Sales'),g('Total Fuel Purchase'),g('Cost of Goods'),g('Gross Profit')
re_,rst,ri,rop,rod=g('Total General'),g('Total Staff'),g('Total Debit Interest'),g('Operating Profit'),g('Owner Debit')
daily=D['daily']
reg=dict(fy='Aug-Nov 2026',hsd=sum(x['hsd'] for x in daily),ms=sum(x['ms'] for x in daily),sales=sum(rs['vals']),purchase=sum(rp['vals']),exp=sum(re_['vals']),staff=sum(rst['vals']),interest=sum(ri['vals']))
reg['gross']=reg['sales']-reg['purchase']; reg['net']=reg['gross']-reg['exp']-reg['staff']-reg['interest']
allr=yrs+[reg]
T={k:sum(x[k] for x in allr) for k in ['hsd','ms','sales','purchase','gross','exp','staff','interest','net']}
oc=sum(v[0] for v in own.values()); od=sum(v[1] for v in own.values())+sum(rod['vals'])
mx=max(x['sales'] for x in allr); mn=max(abs(x['net']) for x in allr)
bars=''.join(f"<div class=br><div class=bl>{x['fy'].replace(' 2026','')}</div><div class=bt><div class=bf style='width:{x['sales']/mx*100:.1f}%'></div></div><div class=bv>{cr(x['sales'])}</div></div>" for x in allr)
nbars=''.join(f"<div class=br><div class=bl>{x['fy'].replace(' 2026','')}</div><div class=bt2><div class=mid></div><div class='bn {'neg' if x['net']<0 else 'pos'}' style='width:{abs(x['net'])/mn*50:.1f}%;{'right:50%' if x['net']<0 else 'left:50%'}'></div></div><div class='bv {'red' if x['net']<0 else 'grn'}'>{cr(x['net'])}</div></div>" for x in allr)
yrow=''.join(f"<tr><td>{x['fy']}{' (Apr-Jul)' if x['fy']=='2026-27' else ''}</td><td>{ind(x['hsd'])}</td><td>{ind(x['ms'])}</td><td>{ind(x['sales'])}</td><td>{ind(x['purchase'])}</td><td>{ind(x['exp']+x['staff'])}</td><td>{ind(x['interest'])}</td><td class={'red' if x['net']<0 else 'grn'}>{ind(x['net'])}</td></tr>" for x in allr)
orow=''.join(f"<tr><td>{k}{' (Apr-Jul)' if k=='2026-27' else ''}</td><td>{ind(c)}</td><td>{ind(d)}</td><td class={'red' if c-d<0 else 'grn'}>{ind(c-d)}</td></tr>" for k,(c,d) in own.items())
mn_=[datetime.date.fromisoformat(m+'-01').strftime('%b-%y') for m in D['reg']['months']]
mrow=''.join(f"<tr><td>{mn_[i]}</td><td>{ind(rs['vals'][i])}</td><td>{ind(rp['vals'][i])}</td><td>{ind(rcg['vals'][i])}</td><td class={'red' if rgp['vals'][i]<0 else 'grn'}>{ind(rgp['vals'][i])}</td><td>{ind(re_['vals'][i]+rst['vals'][i]+ri['vals'][i])}</td><td class={'red' if rop['vals'][i]<0 else 'grn'}>{ind(rop['vals'][i])}</td><td>{ind(rod['vals'][i])}</td></tr>" for i in range(len(mn_)))
# Register data
wb=openpyxl.load_workbook(REG,read_only=True,data_only=True)
pm=[]
for r in wb['Owner & Party Monthly'].iter_rows(min_row=4,values_only=True):
    r=list(r)+[None]*8
    if isinstance(r[0],datetime.datetime) and ((r[3] or 0) or (r[4] or 0)): pm.append((r[0].strftime('%b-%y'),r[3] or 0,r[4] or 0))
fy={'2020-21':4120000,'2021-22':2009500,'2022-23':2590946,'2023-24':2585000,'2024-25':256900,'2025-26':970000,'2026-27 (Apr-Jul)':1337427}
frow=''.join(f"<tr><td>{k}</td><td>{ind(v)}</td><td>-</td></tr>" for k,v in fy.items())+''.join(f"<tr><td>{m} (Register)</td><td>{ind(d)}</td><td>{ind(c)}</td></tr>" for m,c,d in pm)
tot_d=sum(fy.values())+sum(d for _,_,d in pm); tot_c=sum(c for _,c,_ in pm)
jul=[('CHOTU SINGH',350000),('GUDU SINGH',250000),('J P CONST',215975),('Vikash Vishwakrma',28407),('Sonu Singh Ranideva',541290),('Chandan Kumar',698399),('SBI Loan',850000),('Upendra Kumar Singh',1950000),('Rakesh Paswan',92000),('Sharwan Singh',304000),('Rakesh Singh Bairiya',253860),('Shashi Kumar Singh',99000),('Suchun Singh',6069),('Rajan Singh',20516),('Moon Singh',253861),('Dooars Construction',337962),('Ravi Singh',38364),('Rajendra Singh Badiha',20000),('Sonu Badiha',7500),('Ranjan Prajapati',4170)]
now={}
for r in wb['Party Credit Accounts'].iter_rows(min_row=3,values_only=True):
    r=list(r)+[None]*8
    if r[1] and isinstance(r[5],(int,float)) and isinstance(r[0],datetime.datetime): now[str(r[1]).strip().upper()]=(r[5],r[0].strftime('%d-%m-%Y'))
rows=[];seen=set()
for n,j in jul:
    k=n.upper(); seen.add(k); rows.append((n.title() if n.isupper() else n,j,now.get(k)))
for k,v in now.items():
    if k not in seen: rows.append((k.title(),None,v))
prow=''.join(f"<tr><td>{n}</td><td>{ind(j)}</td><td>{ind(v[0]) if v else '-'}</td><td>{v[1] if v else 'Register me nahi'}</td></tr>" for n,j,v in rows)
tj=sum(j for _,j,_ in rows if j); tn=sum(v[0] for _,_,v in rows if v)
# Credit sales party-wise
cs=collections.OrderedDict(); cnote={}
for r in wb['Credit Sales'].iter_rows(min_row=4,max_row=31,values_only=True):
    r=list(r)+[None]*8
    if not r[1]: continue
    k=str(r[1]).strip(); k={'Gufran Kabra':'Gurfan Kabra'}.get(k,k)
    a=cs.setdefault(k,[0,0]); a[0]+=r[3] or 0; a[1]+=r[4] or 0
crow='';CT=[0,0,0]
for k,(bk,pd) in sorted(cs.items(),key=lambda kv:-(kv[1][0]-kv[1][1])):
    f=bk-pd; CT[0]+=bk;CT[1]+=pd;CT[2]+=f
    note='Purana (Mar-2021)' if k in ('Tutu Singh Japla','Roushan Singh Bihra') else ''
    crow+=f"<tr><td>{k}</td><td>{ind(bk)}</td><td>{ind(pd) if pd else '-'}</td><td class={'red' if f>0 else ''}><b>{ind(f)}</b></td><td style='text-align:left'>{note}</td></tr>"
# Nov udhar from DSR
nov=[r for r in daily if r['d'].startswith(TODAY[:7])]
CSS=open('/home/claude/reports/owner.html').read().split('<style>')[1].split('</style>')[0]
CSS+=""".br{display:flex;align-items:center;gap:8px;margin:5px 0;font-size:12px}.bl{width:92px}.bt{flex:1;background:#eef3f9;height:16px;border-radius:4px}.bf{background:#e08a00;height:16px;border-radius:4px}
.bt2{flex:1;background:#eef3f9;height:16px;border-radius:4px;position:relative}.mid{position:absolute;left:50%;top:-2px;bottom:-2px;border-left:1px solid #8a97a8}.bn{position:absolute;top:0;height:16px}.bn.pos{background:#11793a}.bn.neg{background:#c0392b}
.bv{width:100px;text-align:right;font-weight:bold}.cards4{grid-template-columns:repeat(4,1fr)}.c .v{font-size:17px}table{font-size:11.5px}th,td{padding:5px 6px}"""
html=f"""<!doctype html><html><head><meta charset=utf-8><style>{CSS}</style></head><body><div class=page>
<div class=hd><h1>OWNER REPORT - Shuru se ab tak</h1><div class=s>Vanshidhar Petroleum, HPCL Pansa · Shri Ashwini Kumar Singh ji ke liye · Jul-2020 se {TDS} · Pump 1-26 Aug 2026 band tha</div></div>
<h2>Kul hisaab (Jul-2020 se {TDN} 2026)</h2>
<div class="cards cards4" style="margin-top:6px">
<div class=c><div class=l>KUL BIKRI</div><div class=v>{cr(T['sales'])}</div></div>
<div class=c><div class=l>KUL KHARID (FUEL)</div><div class=v>{cr(T['purchase'])}</div></div>
<div class=c><div class=l>HSD (DIESEL)</div><div class=v>{ind(T['hsd'])} L</div></div>
<div class=c><div class=l>MS (PETROL)</div><div class=v>{ind(T['ms'])} L</div></div>
<div class=c><div class=l>PUMP KHARCHA</div><div class=v>{cr(T['exp'])}</div></div>
<div class=c><div class=l>STAFF PAYMENT</div><div class=v>{cr(T['staff'])}</div></div>
<div class=c><div class=l>DEBIT INTEREST</div><div class=v>{cr(T['interest'])}</div></div>
<div class=c><div class=l>NET (KHARID BASIS)</div><div class="v {'red' if T['net']<0 else 'grn'}">{cr(T['net'])}</div></div></div>
<div class=note>Net = Bikri - Kharid - Pump kharcha - Staff - Debit Interest. Tank me bacha stock isme nahi joda gaya, isliye yeh asli munafa/ghata nahi, sirf paisa aane-jaane ka hisaab hai.</div>
<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px"><div><h2>Saal-wise Bikri</h2>{bars}</div><div><h2>Saal-wise Net (kharid basis)</h2>{nbars}</div></div>
<h2>Saal-wise report</h2>
<table><tr><th>Saal (FY)</th><th>HSD (L)</th><th>MS (L)</th><th>Bikri ₹</th><th>Kharid ₹</th><th>Kharcha + Staff ₹</th><th>Interest ₹</th><th>Net ₹</th></tr>{yrow}
<tr class=t><td>TOTAL</td><td>{ind(T['hsd'])}</td><td>{ind(T['ms'])}</td><td>{ind(T['sales'])}</td><td>{ind(T['purchase'])}</td><td>{ind(T['exp']+T['staff'])}</td><td>{ind(T['interest'])}</td><td>{ind(T['net'])}</td></tr></table>
<div class=note>FY = April se March. 2020-21 me Jul-2020 se. Aug-Nov 2026 Register file se (litre 27 Aug se, Daily Sales file se).</div>
<h2>Aug-2026 se mahina-wise (stock jod kar asli Gross Profit)</h2>
<table><tr><th>Mahina</th><th>Bikri ₹</th><th>Kharid ₹</th><th>Maal ki laagat (COGS) ₹</th><th>Gross Profit ₹</th><th>Kharcha+Staff+Interest ₹</th><th>Operating Profit ₹</th><th>Owner Debit ₹</th></tr>{mrow}
<tr class=t><td>Total</td><td>{ind(sum(rs['vals']))}</td><td>{ind(sum(rp['vals']))}</td><td>{ind(sum(rcg['vals']))}</td><td>{ind(sum(rgp['vals']))}</td><td>{ind(sum(re_['vals'])+sum(rst['vals'])+sum(ri['vals']))}</td><td>{ind(sum(rop['vals']))}</td><td>{ind(sum(rod['vals']))}</td></tr></table>
<div class=note>COGS = Opening stock + Kharid - Closing stock (book). Nov {TDN} tak.</div>
<h2>Owner ka len-den (Credit = pump me diya, Debit = pump se liya)</h2>
<table><tr><th>Saal (FY)</th><th>Owner Credit ₹</th><th>Owner Debit ₹</th><th>Net (Cr - Dr) ₹</th></tr>{orow}
<tr><td>Aug-Nov 2026</td><td>-</td><td>{ind(sum(rod['vals']))}</td><td>-</td></tr>
<tr class=t><td>TOTAL</td><td>{ind(oc)}</td><td>{ind(od)}</td><td>{ind(oc-od)}</td></tr></table>
<div class=note>Aug-2026 ke baad Register me owner credit ki koi entry nahi hai, isliye sirf debit dikhaya hai.</div>
<h2>Other Party ka hisaab</h2>
<table><tr><th>Samay</th><th>Party ko diya (Debit) ₹</th><th>Party se mila (Credit) ₹</th></tr>{frow}
<tr class=t><td>TOTAL</td><td>{ind(tot_d)}</td><td>{ind(tot_c)}</td></tr></table>
<div class=note>Jul-2020 se Jul-2026: owner ke account se other party ko gaya paisa (HPCL Reports file). Aug-2026 se: Register ki "Owner &amp; Party Monthly" sheet.</div>
<h2>Party-wise balance ({TDN} tak)</h2>
<table><tr><th>Party</th><th>Balance 31-Jul-2026 ₹</th><th>Abhi ka balance ₹</th><th>Aakhri entry</th></tr>{prow}
<tr class=t><td>TOTAL</td><td>{ind(tj)}</td><td>{ind(tn)}</td><td></td></tr></table>
<div class=note>Abhi ka balance Register ki "Party Credit Accounts" sheet me har party ki aakhri entry se liya gaya hai. Chotu, Gudu aur Moon ke byaj ki shart 3% hai. <b>Dhyan:</b> purani file (File 1) me kul Debit ₹58,53,325 likha hai, lekin party-wise jod ₹17,33,325 aata hai. ₹41,20,000 ka yeh farak CA se milana baaki hai.</div>
<h2>Credit Sale (Udhar) - party-wise baaki ({TDN} tak, Register)</h2>
<table><tr><th>Party</th><th>Udhar diya ₹</th><th>Jama ₹</th><th>Baaki ₹</th><th style="text-align:left">Note</th></tr>{crow}
<tr class=t><td>TOTAL</td><td>{ind(CT[0])}</td><td>{ind(CT[1])}</td><td>{ind(CT[2])}</td><td></td></tr></table>
<div class=note>Kul udhar baaki <b>₹{ind(CT[2])}</b>. Isme Mar-2021 ka purana ₹1,70,000 (Tutu + Roushan) shamil hai. Iske bina chalu udhar <b>₹{ind(CT[2]-170000)}</b> hai.</div>
<h2>{td.strftime('%B %Y')} (1 se {td.day})</h2>
<div class="cards cards4" style="margin-top:6px">
<div class=c><div class=l>BIKRI</div><div class=v>{R(sum(x['total'] for x in nov))}</div></div>
<div class=c><div class=l>MS / HSD</div><div class=v>{ind(sum(x['ms'] for x in nov))} / {ind(sum(x['hsd'] for x in nov))} L</div></div>
<div class=c><div class=l>NAYA UDHAR (DSR)</div><div class=v>{R(sum(x['credit'] for x in nov))}</div></div>
<div class=c><div class=l>SHORT / (EXCESS)</div><div class=v>{R(sum(x['short'] for x in nov))}</div></div></div>
<div class=ft><span>Taiyar: Dharmendra Kumar Singh (Manager)</span><span>Vanshidhar Petroleum, HPCL Pansa</span></div>
</div></body></html>"""
open('/home/claude/reports/owner_all.html','w').write(html)
print('ok',T['sales'],CT,tn)
