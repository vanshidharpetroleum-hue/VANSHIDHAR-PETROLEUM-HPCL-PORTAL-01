import openpyxl,datetime,re,html
F='/home/claude/vanshidhar-petroleum-hpcl-portal-01/Vanshidhar_Petroleum_HPCL_Reports_Jul2020_Jul2026.xlsx'
wb=openpyxl.load_workbook(F,data_only=True)
order=[('Monthly Summery','Mahina-wise hisaab (Bikri, Kharid, Kharcha, Staff, Byaj, Net, Owner)'),
('Sales Detail','Mahina-wise bikri (HSD/MS litre aur rakam)'),
('Purchase Detail','HPCL se tel kharid (invoice-wise, date-wise)'),
('Pump Expense Detail','Pump kharcha (har entry, date-wise)'),
('Staff Pay Detail','Staff payment (mahina-wise, karmchari-wise)'),
('Debit Interest HPCL','HPCL debit interest (invoice-wise)'),
('Bank Interest & Charges','Bank byaj aur charges (date-wise)'),
('Owner Balance Monthly','Owner Credit / Debit / Balance (mahina-wise)'),
('Owner Credit Detail','Owner Credit: pump me aaya paisa (date-wise)'),
('Owner Debit Detail','Owner Debit: nikasi / kharch (date-wise, har entry)'),
('Owner to Other Party','Owner ke account se Other Party ko transfer (date-wise)'),
('Other Party Monthly','Other Party: mahina-wise debit'),
('Other Party Credit Debit','Other Party: party-wise balance (31 Jul 2026)'),
('Party Transfers Oct20-Mar21','Party ledger entries (Oct-2020 se Mar-2021)'),
('Credit Sale Udhar','Credit Sale (Udhar): purana baaki'),
('Mobil Monthly','Mobil / Lube: mahina-wise'),
('Mobil Detail','Mobil / Lube: item-wise detail'),
('Notes & Checks','Notes aur milan (checks)')]
def fmt(v,head=''):
    if v is None: return ''
    if isinstance(v,datetime.datetime):
        return v.strftime('%b-%Y') if (head.lower().startswith(('mahina','period')) and v.day==1) else v.strftime('%d-%m-%Y')
    if isinstance(v,(int,float)):
        if 'fy' in head.lower() or 'invoice' in head.lower() or 'entries' in head.lower(): return str(int(v)) if float(v).is_integer() else str(v)
        neg=v<0; a=abs(v); s='%.2f'%a if (not float(a).is_integer() and ('rate' in head.lower() or a<1000)) else '%.0f'%a
        i,_,d=s.partition('.')
        if len(i)>3: i=re.sub(r'(\d)(?=(\d\d)+$)',r'\1,',i[:-3])+','+i[-3:]
        return ('-' if neg else '')+i+('.'+d if d else '')
    return str(v)
secs=[]
for s,title in order:
    ws=wb[s]; note=ws['A2'].value if ws['A2'].value and len(str(ws['A2'].value))>30 else ''
    hr=None
    for r in range(2,7):
        vals=[c.value for c in ws[r]]
        if sum(1 for x in vals if isinstance(x,str))>=3: hr=r; break
    heads=[c.value for c in ws[hr]]; n=max(i for i,x in enumerate(heads) if x)+1; heads=[str(h or '') for h in heads[:n]]
    rows=[]
    for r in range(hr+1,ws.max_row+1):
        v=[c.value for c in ws[r][:n]]
        if any(x not in (None,'') for x in v): rows.append(v)
    num=[all(isinstance(r[i],(int,float)) or r[i] in (None,'') for r in rows) and any(isinstance(r[i],(int,float)) for r in rows) for i in range(n)]
    h='<h2 id="o%d">%d. %s</h2>'%(len(secs)+1,len(secs)+1,html.escape(title))
    if note: h+='<div class=note>%s</div>'%html.escape(str(note))
    h+='<table><tr>'+''.join('<th class=%s>%s</th>'%('r' if num[i] else 'l',html.escape(heads[i])) for i in range(n))+'</tr>'
    for r in rows:
        bold=any(isinstance(x,str) and re.match(r'(?i)(total|kul|grand)',x) for x in r)
        h+='<tr%s>'%(' class=t' if bold else '')+''.join('<td class=%s>%s</td>'%('r' if num[i] else 'l',html.escape(fmt(r[i],heads[i]))) for i in range(n))+'</tr>'
    secs.append((title,h+'</table>',len(rows)))
toc='<div class=toc><b>Purana hisaab, Jul-2020 se Jul-2026 (File 1 / HPCL Reports file se):</b><ol>'+''.join('<li><a href="#o%d">%s</a> (%d entry)</li>'%(i+1,t,c) for i,(t,_,c) in enumerate(secs))+'</ol></div>'
H='''<!DOCTYPE html><html><head><meta charset=utf-8><style>
body{font-family:Arial,sans-serif;color:#1d2433;margin:0}.page{padding:14px 18px}
.hd{background:#0d4f8b;color:#fff;border-radius:10px;padding:14px 18px}.hd h1{margin:0;font-size:22px}.hd .s{font-size:12px;opacity:.9;margin-top:4px}
h2{color:#0d4f8b;font-size:15px;border-bottom:2px solid #dfe6ef;padding-bottom:4px;margin:20px 0 6px;break-after:avoid}
table{width:100%;border-collapse:collapse;font-size:8.5px}th{background:#eef3f9}th,td{border:1px solid #d5dde8;padding:2px 4px}
tr{break-inside:avoid}.l{text-align:left}.r{text-align:right;white-space:nowrap}tr.t td{font-weight:bold;background:#f4f7fb}
.note{font-size:10px;color:#556;margin:2px 0 6px}.toc{font-size:12px;margin-top:12px;line-height:1.6}.toc a{color:#0d4f8b}
</style></head><body><div class=page><div class=hd><h1>PURANA HISAAB - Jul-2020 se Jul-2026</h1><div class=s>Vanshidhar Petroleum, HPCL Pansa · Final Closing (File 1) se · har mahina aur har entry</div></div>'''+toc+''.join(h for _,h,_ in secs)+'</div></body></html>'
open('old.html','w').write(H); print([(t,c) for t,_,c in secs])
