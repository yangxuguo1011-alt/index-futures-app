# -*- coding: utf-8 -*-
"""后台定时服务（python-for-android Service）。

APP 退出后由系统托管，定时检查是否需要刷新并执行一次分析。
结果与 APP 写在同一目录，下次打开界面直接显示。

buildozer.spec 需声明：services = Tracker:service.py
（对应 Java 类 org.example.indexfuturestracker.ServiceTracker）

第一次运行会尝试把自己提升为前台服务并挂常驻通知（安卓 8+ 必需），
否则系统在几分钟内就会把后台服务杀掉。若提升失败则退化为普通后台服务继续跑。
"""

import time
import traceback

import app_job

CHECK_INTERVAL = 15 * 60   # 每 15 分钟检查一次是否需要刷新


def _try_foreground():
    """安卓 8+ 后台服务会被系统回收，必须挂前台通知才活得久。失败不影响主逻辑。"""
    try:
        from jnius import autoclass
        svc = autoclass("org.kivy.android.PythonService").mService
        SDK_INT = autoclass("android.os.Build$VERSION").SDK_INT
        NBuilder = autoclass("android.app.Notification$Builder")
        if SDK_INT >= 26:
            NChannel = autoclass("android.app.NotificationChannel")
            NMgr = autoclass("android.app.NotificationManager")
            chan = NChannel("tracker", "持仓追踪", NMgr.IMPORTANCE_LOW)
            svc.getSystemService(NMgr).createNotificationChannel(chan)
            nb = NBuilder(svc, "tracker")
        else:
            nb = NBuilder(svc)
        nb.setContentTitle("股指期货持仓追踪")
        nb.setContentText("定时抓取运行中")
        nb.setSmallIcon(svc.getApplicationInfo().icon)
        svc.startForeground(1, nb.build())
        return True
    except Exception:
        traceback.print_exc()
        return False


def main():
    app_job.setup_storage()
    _try_foreground()
    while True:
        try:
            if app_job.should_refresh():
                app_job.run()
        except Exception:
            traceback.print_exc()
        time.sleep(CHECK_INTERVAL)


if __name__ == "__main__":
    main()
