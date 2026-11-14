import os,base64
from playwright.sync_api import sync_playwright
from pypdf import PdfWriter,PdfReader
ROOT='/home/claude/vanshidhar-petroleum-hpcl-portal-01'
tabs=[('total','Dashboard: Total (Jul-2020 se ab tak)'),('daily','Dashboard: Roz ki Bikri'),('month','Dashboard: Mahina-wise (2020-26)'),('year','Dashboard: Saal-wise (FY)'),('reg','Dashboard: Register Aug-2026 se')]
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':1280,'height':900},device_scale_factor=1.5)
    pg.route('**/chart.umd.min.js',lambda r:r.fulfill(path='/tmp/package/dist/chart.umd.js',content_type='application/javascript'))
    pg.goto('file://'+ROOT+'/dashboard.html'); pg.wait_for_timeout(1200)
    pg.add_style_tag(content='header,nav{position:static!important}')
    for t,_ in tabs:
        pg.click('button[data-t="%s"]'%t); pg.wait_for_timeout(900)
        pg.screenshot(path='tab_%s.png'%t,full_page=True)
    sec=''
    for t,title in tabs:
        img=base64.b64encode(open('tab_%s.png'%t,'rb').read()).decode()
        sec+='<section><h2>%s</h2><img src="data:image/png;base64,%s"></section>'%(title,img)
    cover='''<section class=cover><div class=box><div class=t1>VANSHIDHAR PETROLEUM</div><div class=t2>HPCL Pansa</div><div class=t3>OWNER REPORT</div>
<div class=t4>Shuru (Jul-2020) se aaj (14-11-2026) tak<br>Summary · Dashboard aur Chart · Poora date-wise hisaab</div>
<div class=t5>Shri Ashwini Kumar Singh ji (Owner) ke liye<br>Taiyar: Dharmendra Kumar Singh (Manager)</div></div>
<div class=toc><b>Is file me:</b><ol><li>Dashboard aur Chart (5 hisse)</li><li>Owner Report Summary (Jul-2020 se ab tak)</li><li>Poora date-wise hisaab (1 Aug 2026 se 14 Nov 2026, 13 hisse)</li></ol></div></section>'''
    H='''<!DOCTYPE html><html><head><meta charset=utf-8><style>body{margin:0;font-family:Arial,sans-serif}section{page-break-after:always;padding:10px 16px}
h2{color:#0d4f8b;font-size:16px;margin:4px 0 8px}img{width:100%%}.cover{height:270mm;display:flex;flex-direction:column;justify-content:center}
.box{background:#0d4f8b;color:#fff;border-radius:14px;padding:40px;text-align:center}.t1{font-size:34px;font-weight:bold}.t2{font-size:20px;margin-top:6px}
.t3{font-size:28px;margin-top:30px;color:#ffd34a;font-weight:bold}.t4{font-size:16px;margin-top:14px;line-height:1.6}.t5{font-size:14px;margin-top:30px;opacity:.9;line-height:1.6}
.toc{margin-top:30px;font-size:15px;line-height:1.8}</style></head><body>%s%s</body></html>'''%(cover,sec)
    open('dash.html','w').write(H)
    pg2=b.new_page(); pg2.goto('file:///home/claude/combo/dash.html'); pg2.wait_for_timeout(500)
    pg2.pdf(path='dash.pdf',format='A4',print_background=True,margin={'top':'6mm','bottom':'6mm','left':'0','right':'0'})
    b.close()
w=PdfWriter()
for f in ['dash.pdf','/home/claude/reports/Owner_Report_14Nov2026.pdf','/home/claude/full/Owner_Report_Full_DateWise_14Nov2026.pdf']:
    for pgx in PdfReader(f).pages: w.add_page(pgx)
out='Vanshidhar_Petroleum_HPCL_Pansa_Owner_Report_Complete_14Nov2026.pdf'
w.write(out); print(len(w.pages))
