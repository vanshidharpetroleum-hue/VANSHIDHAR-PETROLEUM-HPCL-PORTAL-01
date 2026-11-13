"""Mobile lock screen ke liye lockscreen.png (1080x2400) banata hai. build_dashboard.py ke baad chalayein."""
import os
from playwright.sync_api import sync_playwright
ROOT=os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
local=[x for x in ['/tmp/package/dist/chart.umd.js'] if os.path.exists(x)]
with sync_playwright() as p:
    b=p.chromium.launch(); pg=b.new_page(viewport={'width':432,'height':960},device_scale_factor=2.5)
    if local: pg.route('**/chart.umd.min.js',lambda r:r.fulfill(path=local[0],content_type='application/javascript'))
    pg.goto('file://'+os.path.join(ROOT,'wallpaper.html')+'?lock=1'); pg.wait_for_timeout(600)
    pg.screenshot(path=os.path.join(ROOT,'lockscreen.png')); b.close()
print('lockscreen.png')
