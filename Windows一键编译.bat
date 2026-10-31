@echo off
chcp 65001 >nul
cd /d "%~dp0"
title APK 一键编译（Windows 10 / WSL）

echo ============================================================
echo   股指期货持仓追踪 · APK 一键编译
echo   适用：Windows 10（2004 以上版本）
echo ============================================================
echo.

REM ---- check if a working WSL distro exists ----
wsl bash -lc "echo WSL_OK" > "%TEMP%\wsltest.txt" 2>nul
find "WSL_OK" "%TEMP%\wsltest.txt" >nul
if not errorlevel 1 goto BUILD

REM ================= NO WSL YET -> INSTALL =================
echo ------------------------------------------------------------
echo   检测到这台电脑还没装 Linux 子系统（WSL），现在开始安装。
echo   接下来会弹出一个蓝色/黑色的窗口，可能会要求你点“是”。
echo ------------------------------------------------------------
echo.
echo   【先确认一件小事】
echo     按 Win 键 + R，输入 winver 回车，看版本号：
echo     必须是 19041 或更大（也就是 2004 及以上）。
echo     如果比这个小，本方案用不了，请改用 Docker 或 GitHub 方式。
echo.
pause

echo.
echo 正在安装 WSL + Ubuntu（约 5-15 分钟，取决于网速）...
powershell -Command "Start-Process -FilePath 'wsl.exe' -ArgumentList '--install','-d','Ubuntu' -Verb RunAs -Wait"

echo.
echo ============================================================
echo   安装命令已执行完毕。
echo.
echo   接下来请做两件事：
echo.
echo   1) 重启电脑。
echo   2) 重启后会自动弹出一个 Ubuntu 黑窗口，
echo      让你输入用户名和密码 —— 自己设一个记住（密码不显示，
echo      输完直接回车）。设完就可以把那个窗口关掉。
echo.
echo   然后【再次双击本脚本】，就会自动开始编译 APK。
echo ============================================================
echo.
pause
exit /b 0

REM ================= WSL READY -> BUILD =================
:BUILD
echo 检测到 WSL 已就绪，开始准备编译环境。
echo.

REM convert current Windows path to a WSL path
wsl wslpath -u "%CD%" > "%TEMP%\wslpath.txt" 2>nul
set /p WSLPATH=<"%TEMP%\wslpath.txt"
if not defined WSLPATH (
    echo 路径转换失败，请把整个文件夹放到不含中文的路径下再试。
    pause
    exit /b 1
)
echo   工程路径：%WSLPATH%
echo.

echo 第 1 步：把代码复制进 Linux 子系统...
wsl bash -lc "rm -rf ~/android_app && cp -r '%WSLPATH%' ~/android_app && echo COPY_OK"

echo.
echo 第 2 步：开始编译（首次约 20-40 分钟，要下载 1.5GB 工具，
echo                     建议保持梯子开启，中途不要关窗口）
echo.
wsl bash -lc "cd ~/android_app && chmod +x build_apk.sh && ./build_apk.sh arm64"

echo.
echo 第 3 步：把 APK 取回到 Windows...
if not exist bin mkdir bin
wsl bash -lc "cp ~/android_app/bin/*.apk '%WSLPATH%/bin/' 2>/dev/null; ls -lh ~/android_app/bin/*.apk 2>/dev/null || echo NO_APK"

echo.
echo ============================================================
if exist bin\*.apk (
    echo   成功！APK 已经放在本目录下的 bin 文件夹里：
    echo     %CD%\bin
    echo.
    echo   把它用微信发给自己，或数据线拷到手机，点开安装即可。
) else (
    echo   没有生成 APK，说明编译失败了。
    echo.
    echo   请把上面窗口里的红色报错信息截图发给你做程序的人。
    echo   也可以执行下面这条命令把日志存成文件：
    echo     wsl bash -lc "cd ~/android_app && buildozer -v android debug 2>&1 | tee build.log"
    echo   日志在：%CD%\build.log
)
echo ============================================================
echo.
pause
