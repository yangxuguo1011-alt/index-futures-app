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

# Android 构建配置（p4a v2026.05.09 推荐 NDK 28c，最低 25；target API 34）
# minapi 24：Python 3.14 用到 preadv/pwritev，这两个函数 Android API 24 才声明，
# 保持 23 会在编译 python3 recipe 时报 "undeclared function"
android.api = 34
android.minapi = 24
android.ndk = 28c
android.archs = arm64-v8a
android.enable_androidx = True
android.allow_backup = True
android.accept_sdk_license = True

# p4a 用稳定发行版（master 的 libffi 3.4.6 配方在 CI 上 autoconf 会报
# LT_SYS_SYMBOL_USCORE 未定义；v2026.05.09 的 libffi 3.4.2 配方稳定）
p4a.branch = v2026.05.09
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
