[app]

# 应用信息（安装后显示的名字 / 包名，包名决定数据目录 /data/data/<pkg>）
title = 股指期货持仓追踪
package.name = indexfuturestracker
package.domain = org.example

source.dir = .
source.include_exts = py,png,jpg,kv,atlas,json,ttf,otf,html,txt
source.exclude_dirs = tests,bin,build,.venv,data,outputs,盘面多维分析,.github
source.exclude_exts = bak,log

version = 1.1
requirements = python3,kivy,requests,openpyxl,pyjnius,android

# 后台定时服务（对应 service.py，Java 类名 ServiceTracker）
services = Tracker:service.py

orientation = portrait
fullscreen = 0
android.permissions = INTERNET,ACCESS_NETWORK_STATE,FOREGROUND_SERVICE,POST_NOTIFICATIONS,WAKE_LOCK

# Android 构建配置（34 + NDK 25b 是目前 p4a 上最稳的组合；
# API 35 需要 NDK r27+，不建议轻易改，改了容易编不过）
android.api = 34
android.minapi = 23
android.ndk = 25b
android.archs = arm64-v8a,armeabi-v7a
android.enable_androidx = True
android.allow_backup = True
android.accept_sdk_license = True

# p4a 用 master 分支（含较新的 NDK/API 修复）
p4a.branch = master
p4a.hook =
p4a.port =

# 图标（可选：放 icon.png 到工程根目录才会生效）
#icon.filename = icon.png
#presplash.filename = presplash.png

[buildozer]
log_level = 2
warn_on_root = 1
build_dir = ./.buildozer
bin_dir = ./bin

[buildozer:app:android.gradle]

[buildozer:app:android]
