"""Excel files se dashboard.html banata hai.
Chalane ka tarika (repo folder me):  python tools/build_dashboard.py
"""
import json, datetime, os, openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
def p(n): return os.path.join(ROOT, n)
def num(v):
    return round(float(v), 2) if isinstance(v, (int, float)) else 0.0
import shutil, subprocess, tempfile
_TMP = tempfile.mkdtemp()
def wb(n):
    # Excel formulas ko LibreOffice se dobara calculate karke padhta hai (agar LibreOffice hai)
    out = os.path.join(_TMP, 'out', os.path.basename(n))
    if not os.path.exists(out) and shutil.which('soffice'):
        prof = os.path.join(_TMP, 'prof', 'user'); os.makedirs(prof, exist_ok=True)
        open(os.path.join(prof, 'registrymodifications.xcu'), 'w').write(
            '<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry">'
            '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse">'
            '<value>0</value></prop></item></oor:items>')
        subprocess.run(['soffice', '-env:UserInstallation=file://' + os.path.join(_TMP, 'prof'), '--headless', '--convert-to', 'xlsx', '--outdir', os.path.join(_TMP, 'out'), p(n)],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return openpyxl.load_workbook(out if os.path.exists(out) else p(n), read_only=True, data_only=True)

# 1. Daily sales (Aug-2026 se)
daily = []
ws = wb('Vanshidhar Daily_Sales_Report_Aug2026_to.xlsx')['Daily Summery']
for r in ws.iter_rows(min_row=5, values_only=True):
    if not r or not isinstance(r[0], datetime.datetime): continue
    r = list(r) + [None]*20
    total = num(r[6])
    if total <= 0: continue
    daily.append(dict(d=r[0].strftime('%Y-%m-%d'), ms=num(r[1]), hsd=num(r[2]), lube=num(r[5]), total=total,
        cash=num(r[7]), upi=num(r[8]), owner=num(r[9]), party=num(r[10]), credit=num(r[11]),
        short=num(r[13]), kharcha=num(r[14])))

# 2. Monthly Jul-2020 to Jul-2026
monthly = []
ws = wb('Vanshidhar_Petroleum_HPCL_Reports_Jul2020_Jul2026.xlsx')['Monthly Summery']
for r in ws.iter_rows(min_row=5, values_only=True):
    if not r or not isinstance(r[0], datetime.datetime): continue
    r = list(r) + [None]*20
    monthly.append(dict(m=r[0].strftime('%Y-%m'), fy=r[1], hsd=num(r[2]), ms=num(r[3]), sales=num(r[4]),
        purchase=num(r[5]), gross=num(r[6]), exp=num(r[7]), staff=num(r[8]), interest=num(r[9]), net=num(r[10])))

# 3. Yearly
yearly = []
ws = wb('Vanshidhar_Petroleum_HPCL_Reports_Jul2020_Jul2026.xlsx')['Yearly Report']
for r in ws.iter_rows(values_only=True):
    if r and isinstance(r[0], str) and (r[0][:2] == '20' or r[0] == 'TOTAL') and isinstance(r[1], (int, float)):
        yearly.append(dict(fy=r[0], hsd=num(r[1]), ms=num(r[2]), sales=num(r[3]), purchase=num(r[4]),
            gross=num(r[5]), exp=num(r[6]), staff=num(r[7]), interest=num(r[8]), net=num(r[9])))

# 4. Register Aug-2026 se (monthly summary)
reg = {'months': [], 'rows': []}
ws = wb('Vanshidhar Register_Aug-2026 to _Complete.xlsx')['Monthly Summery']
rows = list(ws.iter_rows(values_only=True))
hdr = rows[1]
idx = [i for i, c in enumerate(hdr) if isinstance(c, datetime.datetime)]
tot_i = [i for i, c in enumerate(hdr) if c == 'Total']
for r in rows[2:]:
    if r and isinstance(r[0], str) and r[0].strip():
        vals = [num(r[i]) for i in idx]
        if not any(vals): continue
        reg['rows'].append(dict(name=r[0].strip(), vals=vals, total=num(r[tot_i[0]]) if tot_i else None))
months_used = [i for k, i in enumerate(idx) if any(row['vals'][k] for row in reg['rows'])]
keep = [idx.index(i) for i in months_used]
reg['months'] = [hdr[i].strftime('%Y-%m') for i in months_used]
for row in reg['rows']: row['vals'] = [row['vals'][k] for k in keep]

desh = {}
for r in wb('Vanshidhar Register_Aug-2026 to _Complete.xlsx')['Deshbord'].iter_rows(min_row=3, values_only=True):
    if r and isinstance(r[0], str) and isinstance(r[1], (int, float)): desh[r[0].strip()] = round(r[1], 2)
data = dict(desh=desh, built=datetime.datetime.now().strftime('%d-%m-%Y %H:%M'), daily=daily, monthly=monthly, yearly=yearly, reg=reg)
tpl = open(p('tools/dashboard_template.html'), encoding='utf-8').read()
open(p('dashboard.html'), 'w', encoding='utf-8').write(tpl.replace('/*DATA*/null', json.dumps(data, ensure_ascii=False)))
wtpl = open(p('tools/wallpaper_template.html'), encoding='utf-8').read()
open(p('wallpaper.html'), 'w', encoding='utf-8').write(wtpl.replace('/*DATA*/null', json.dumps(data, ensure_ascii=False)))
print('dashboard.html bana:', len(daily), 'din,', len(monthly), 'mahine,', len(yearly), 'saal')
