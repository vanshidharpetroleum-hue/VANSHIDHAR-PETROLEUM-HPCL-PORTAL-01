@echo off
REM Vanshidhar Petroleum - har 30 minute me naya dashboard photo laata hai aur DESKTOP WALLPAPER par laga deta hai.
REM Photo: C:\PumpLockScreen\pump_lockscreen.png  (Lock screen Slideshow ke liye bhi yahi folder chuniye)
if not exist "C:\PumpLockScreen" mkdir "C:\PumpLockScreen"
start "" /min powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command "$sig='[DllImport(\"user32.dll\",CharSet=CharSet.Auto)] public static extern int SystemParametersInfo(int a,int b,string c,int d);'; Add-Type -MemberDefinition $sig -Name W -Namespace P; while($true){ try { $f='C:\PumpLockScreen\pump_lockscreen.png'; Invoke-WebRequest -UseBasicParsing -Uri ('https://vanshidharpetroleum-hue.github.io/VANSHIDHAR-PETROLEUM-HPCL-PORTAL-01/LockScreen/laptop_lockscreen.png?t=' + [DateTimeOffset]::Now.ToUnixTimeSeconds()) -OutFile $f; [P.W]::SystemParametersInfo(20,0,$f,3) | Out-Null } catch {}; Start-Sleep -Seconds 1800 }"
