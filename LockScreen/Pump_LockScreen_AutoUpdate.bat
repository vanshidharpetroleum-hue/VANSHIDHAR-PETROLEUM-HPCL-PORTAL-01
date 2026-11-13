@echo off
REM Vanshidhar Petroleum - Laptop lock screen photo har 30 minute me GitHub se naya laata hai.
REM Is file ko Startup folder me rakhiye (Windows+R > shell:startup).
if not exist "C:\PumpLockScreen" mkdir "C:\PumpLockScreen"
start "" /min powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command "while($true){ try { Invoke-WebRequest -UseBasicParsing -Uri ('https://vanshidharpetroleum-hue.github.io/VANSHIDHAR-PETROLEUM-HPCL-PORTAL-01/LockScreen/laptop_lockscreen.png?t=' + [DateTimeOffset]::Now.ToUnixTimeSeconds()) -OutFile 'C:\PumpLockScreen\pump_lockscreen.png' } catch {}; Start-Sleep -Seconds 1800 }"
