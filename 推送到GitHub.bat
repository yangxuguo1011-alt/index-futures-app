@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo   Step 1 of 3 : upload code to GitHub
echo   (turn on your VPN/proxy first if github.com is blocked)
echo ============================================================
echo.
echo  Before running this you must create an EMPTY repository on
echo  github.com. Do NOT add README / .gitignore / license.
echo.

set /p REPO= Paste repo URL here (https://github.com/xxx/yyy.git) :

if "%REPO%"=="" (
    echo.
    echo  No URL entered. Nothing was done.
    pause
    exit /b 1
)

echo.
echo  If you are using a VPN/proxy, enter its port number.
echo  Common values:  Clash / Clash Verge = 7890
echo                  v2rayN             = 10809
echo                  Shadowsocks        = 1080
echo  If you have no proxy, just press Enter.
echo.
set /p PORT= Proxy port (or Enter to skip) :

if not "%PORT%"=="" (
    git config --local http.proxy http://127.0.0.1:%PORT%
    git config --local https.proxy http://127.0.0.1:%PORT%
    echo  Proxy set to 127.0.0.1:%PORT%
) else (
    git config --local --unset http.proxy  >nul 2>&1
    git config --local --unset https.proxy >nul 2>&1
    echo  No proxy configured.
)

git remote remove origin >nul 2>&1
git remote add origin "%REPO%"
git branch -M main
git add -A
git commit -m "update" >nul 2>&1

echo.
echo  Uploading... A GitHub sign-in window may pop up - please sign in.
echo  (If you use a proxy browser login, choose "Sign in with your browser")
echo.
git push -u origin main

echo.
if errorlevel 1 (
    echo ============================================================
    echo   UPLOAD FAILED - nothing is lost, you can just run again.
    echo   Usual reasons:
    echo     1. Repository URL is wrong
    echo     2. Not signed in / sign-in window was ignored
    echo     3. Repository was not created as EMPTY
    echo     4. Proxy port is wrong (rerun and try another port)
    echo ============================================================
) else (
    echo ============================================================
    echo   UPLOADED OK!
    echo.
    echo   Next: go to github.com, open your repository,
    echo   click the "Actions" tab. A build starts by itself.
    echo   Wait 20-40 minutes, then download the APK from
    echo   "Artifacts" at the bottom of that run page.
    echo ============================================================
)
echo.
pause
