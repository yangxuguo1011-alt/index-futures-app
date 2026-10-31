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

# Android 构建配置（target API 34 + NDK 25b：p4a 2024 稳定组合）
# minapi 24（Android 7.0+）
android.api = 34
android.minapi = 24
android.ndk = 25b
android.archs = arm64-v8a
android.enable_androidx = True
android.allow_backup = True
android.accept_sdk_license = True

# p4a 用 2024.01.21 稳定版（Python 3.11 + 老版 pip）：
#   master 的 libffi 3.4.6 配方 autoconf 会报 LT_SYS_SYMBOL_USCORE 未定义；
#   v2026.05.09 用 Python 3.14/pip 26，会拒绝 android 平台标签的 wheel
#   （charset_normalizer ... android_24_arm64_v8a.whl is not a supported wheel）。
p4a.branch = v2024.01.21

# 注意：不要保留空的 `p4a.hook =` / `p4a.port =`，
# buildozer 会把空值当路径传 `--hook=<工程目录>`，p4a 加载 hook 模块时
# 报 AttributeError: 'NoneType' object has no attribute 'loader'（打包最后一步失败）。

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
