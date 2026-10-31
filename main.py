# -*- coding: utf-8 -*-
"""股指期货持仓追踪 · 安卓版主界面（Kivy）。

界面只做两件事：展示 app_job 产出的分析，以及触发刷新。
所有抓取/计算/结论逻辑都在 daily_job.py 里，与电脑端完全同源。

刷新策略（安卓后台不可靠，做了三层）：
  1. 启动即检查：本地没有当日完整数据就自动跑一次；
  2. 前台轮询：APP 开着时每 10 分钟检查一次是否需要刷新；
  3. 后台服务：p4a Service（service.py），APP 退出后定时跑，结果写入同一目录，
     下次打开直接显示。
"""

import os
import threading
import datetime
import traceback

from kivy.app import App
from kivy.clock import Clock, mainthread
from kivy.core.text import LabelBase
from kivy.properties import StringProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.vector import Vector

import app_job

# 全局字体缩放系数（双指捏合调整，render 时生效）
ZOOM = [1.0]
ZOOM_MIN, ZOOM_MAX = 0.6, 2.5

# ---------------- 中文字体（Kivy 默认 Roboto 不含中文，会显示方块）----------------
# 优先用打包进 assets 的字体；其次用系统自带的中文字体（各机型路径不同，逐个试）。
CJK_CANDIDATES = [
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets",
                 "NotoSansSC-Regular.ttf"),
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets",
                 "NotoSansSC-Regular.otf"),
    "/system/fonts/NotoSansCJK-Regular.ttc",
    "/system/fonts/NotoSansSC-Regular.otf",
    "/system/fonts/DroidSansFallback.ttf",
    "/system/fonts/SourceHanSansCN-Regular.otf",
    "/system/fonts/Miui-Regular.ttf",
    "/system/fonts/HuaweiSans-Regular.ttf",
]


def setup_font():
    """把 Kivy 默认字体替换成中文字体。失败则返回 None（此时中文可能显示为方块）。"""
    for p in CJK_CANDIDATES:
        if not os.path.exists(p):
            continue
        try:
            LabelBase.register(name="Roboto", fn_regular=p)
            return p
        except Exception:
            continue
    return None


FONT_PATH = setup_font()

C_TEXT = (0.13, 0.13, 0.13, 1)
C_MUTE = (0.42, 0.42, 0.42, 1)
C_BLUE = (0.13, 0.40, 0.80, 1)
C_RED = (0.80, 0.18, 0.15, 1)      # 涨/净空扩大
C_GREEN = (0.09, 0.50, 0.22, 1)    # 跌（A股习惯）
C_BG = (0.96, 0.96, 0.97, 1)


def _bind_wrap(w):
    w.bind(size=lambda *_a: setattr(w, "text_size", (w.width - 4, None)))
    w.bind(texture_size=lambda *_a: setattr(w, "height", max(w.texture_size[1], 24)))
    return w


def para(text, size=14, color=C_TEXT, bold=False):
    l = Label(text=text, font_size=size * ZOOM[0], color=color, bold=bold,
              size_hint_y=None, halign="left", valign="top")
    return _bind_wrap(l)


def section(parent, text):
    l = Label(text=text, font_size=15 * ZOOM[0], bold=True, color=C_BLUE,
              size_hint_y=None, height=int(36 * ZOOM[0]), halign="left", valign="middle")
    l.bind(size=lambda *_a: setattr(l, "text_size", (l.width - 4, None)))
    parent.add_widget(l)


def kv(parent, k, v, size=14, vcolor=C_TEXT):
    row = BoxLayout(orientation="horizontal", size_hint_y=None,
                    height=int(size * 2.4 * ZOOM[0]), padding=(0, 0))
    a = Label(text=k, font_size=size * ZOOM[0], color=C_MUTE, size_hint_x=0.5, halign="left")
    b = Label(text=str(v), font_size=size * ZOOM[0], bold=True, color=vcolor,
              size_hint_x=0.5, halign="right")
    a.bind(size=lambda *_a: setattr(a, "text_size", (a.width - 4, None)))
    b.bind(size=lambda *_a: setattr(b, "text_size", (b.width - 4, None)))
    row.add_widget(a)
    row.add_widget(b)
    parent.add_widget(row)


def card(parent, text, size=14, color=C_TEXT):
    box = BoxLayout(orientation="vertical", size_hint_y=None, padding=(10, 8))
    l = para(text, size=size, color=color)
    box.add_widget(l)
    box.bind(minimum_height=box.setter("height"))
    parent.add_widget(box)


class Root(BoxLayout):
    status = StringProperty("就绪")

    def __init__(self, **kw):
        super(Root, self).__init__(orientation="vertical", **kw)
        self.busy = False
        self.auto = True
        self.zoom = 1.0                  # 当前缩放倍率（双指捏合调整）
        self._pinch_touches = []         # 参与捏合的触点
        self._pinch_start_dist = None    # 捏合起始时两指间距
        self._pinch_start_zoom = 1.0
        self._zoom_apply = None          # 节流用的 Clock 句柄
        self._last_result = None         # 最近一次渲染的数据（缩放后重渲染用）
        self._last_full = datetime.datetime.min  # 上次完整刷新时间（限流）
        self._build()
        Clock.schedule_once(self._boot, 0.2)

    # ---------- 双指捏合缩放 ----------
    def on_touch_down(self, touch):
        self._pinch_touches.append(touch)
        if len(self._pinch_touches) == 2:
            t1, t2 = self._pinch_touches[:2]
            self._pinch_start_dist = Vector(t1.pos).distance(t2.pos)
            self._pinch_start_zoom = self.zoom
        return super(Root, self).on_touch_down(touch)

    def on_touch_move(self, touch):
        if (self._pinch_start_dist and len(self._pinch_touches) >= 2
                and touch in self._pinch_touches):
            t1, t2 = self._pinch_touches[:2]
            d = Vector(t1.pos).distance(t2.pos)
            if d > 10:
                new_zoom = max(ZOOM_MIN, min(ZOOM_MAX,
                              self._pinch_start_zoom * d / self._pinch_start_dist))
                if abs(new_zoom - self.zoom) >= 0.02:
                    self.zoom = new_zoom
                    self._schedule_zoom()
        return super(Root, self).on_touch_move(touch)

    def on_touch_up(self, touch):
        if touch in self._pinch_touches:
            self._pinch_touches.remove(touch)
            if len(self._pinch_touches) < 2:
                self._pinch_start_dist = None
        return super(Root, self).on_touch_up(touch)

    def _schedule_zoom(self):
        """缩放手势期间节流重渲染（约 20fps），避免卡顿。"""
        if self._zoom_apply is not None:
            Clock.unschedule(self._zoom_apply)
        self._zoom_apply = Clock.schedule_once(self._apply_zoom, 0.05)

    def _apply_zoom(self, *_a):
        self._zoom_apply = None
        ZOOM[0] = self.zoom
        if self._last_result:
            self.render(self._last_result)

    # ---------- 布局 ----------
    def _build(self):
        head = BoxLayout(orientation="vertical", size_hint_y=None, height=64,
                         padding=(14, 6))
        t = Label(text="股指期货持仓追踪", font_size=19, bold=True, color=C_TEXT,
                  size_hint_y=None, height=30, halign="left")
        t.bind(size=lambda *_a: setattr(t, "text_size", (t.width - 4, None)))
        self.sub = Label(text="", font_size=12, color=C_MUTE,
                         size_hint_y=None, height=22, halign="left")
        self.sub.bind(size=lambda *_a: setattr(self.sub, "text_size",
                                               (self.sub.width - 4, None)))
        head.add_widget(t)
        head.add_widget(self.sub)
        self.add_widget(head)

        btns = BoxLayout(orientation="horizontal", size_hint_y=None, height=52,
                         spacing=8, padding=(12, 6))
        self.btn_ref = Label(text="[ 立即刷新 ]", font_size=15, bold=True,
                             color=C_BLUE, size_hint_x=0.5)
        self.btn_ref.bind(on_touch_down=lambda w, m: (self.refresh()
                                                      if w.collide_point(*m.pos) else None))
        self.btn_auto = Label(text="[ 自动刷新：开 ]", font_size=14,
                              color=C_GREEN, size_hint_x=0.5)
        self.btn_auto.bind(on_touch_down=lambda w, m: (self.toggle_auto()
                                                       if w.collide_point(*m.pos) else None))
        btns.add_widget(self.btn_ref)
        btns.add_widget(self.btn_auto)
        self.add_widget(btns)

        self.st = Label(text="", font_size=12, color=C_MUTE, size_hint_y=None, height=26,
                        halign="left")
        self.st.bind(size=lambda *_a: setattr(self.st, "text_size", (self.st.width - 4, None)))
        self.add_widget(self.st)

        sc = ScrollView(size_hint=(1, 1))
        self.content = BoxLayout(orientation="vertical", size_hint_y=None,
                                 spacing=2, padding=(14, 8, 14, 20))
        self.content.bind(minimum_height=self.content.setter("height"))
        sc.add_widget(self.content)
        self.add_widget(sc)

    # ---------- 生命周期 ----------
    def _boot(self, *_a):
        app_job.setup_storage()
        self.render(app_job.load_cached())
        if app_job.should_refresh():
            self._last_full = datetime.datetime.now()
            self.refresh()
        else:
            self.set_status("本地已有最新数据")
        Clock.schedule_interval(self._tick, 60)  # 每分钟：盘中轻量刷新 + 完整刷新检查

    def _tick(self, *_a):
        if not self.auto or self.busy:
            return
        # 盘中轻量实时刷新：只拉行情/宽度/资金/隐波基差，不重跑持仓分析
        try:
            r = app_job.refresh_quotes()
            if r:
                self.render(r)
                self.set_status("实时刷新 %s" % datetime.datetime.now().strftime("%H:%M:%S"))
                return
        except Exception:
            traceback.print_exc()
        # 完整刷新（如收盘后抓当日持仓），最小间隔 10 分钟，避免反复重跑
        now = datetime.datetime.now()
        if app_job.should_refresh() and (now - self._last_full).total_seconds() >= 600:
            self._last_full = now
            self.refresh()

    def toggle_auto(self):
        self.auto = not self.auto
        self.btn_auto.text = "[ 自动刷新：%s ]" % ("开" if self.auto else "关")
        self.btn_auto.color = C_GREEN if self.auto else C_MUTE

    def refresh(self):
        if self.busy:
            return
        self.busy = True
        self.btn_ref.color = C_MUTE
        threading.Thread(target=self._work, daemon=True).start()

    def _work(self):
        try:
            r = app_job.run(progress=self.set_status)
            self._done(r, None)
        except Exception as e:
            traceback.print_exc()
            self._done(None, e)

    @mainthread
    def _done(self, r, err):
        self.busy = False
        self.btn_ref.color = C_BLUE
        self._last_full = datetime.datetime.now()
        if err is not None:
            self.status = "刷新失败：%s" % err
            self.st.text = "刷新失败：%s" % err
            return
        self.render(r)
        self.set_status("已更新 %s" % datetime.datetime.now().strftime("%H:%M"))

    @mainthread
    def set_status(self, msg):
        self.st.text = str(msg)

    # ---------- 渲染 ----------
    def render(self, r):
        self._last_result = r   # 缩放手势需要用最近数据重渲染
        c = self.content
        c.clear_widgets()
        if not r:
            self.sub.text = "尚无数据"
            card(c, "还没有分析数据。点上方「立即刷新」开始抓取（需联网，约 1 分钟）。")
            return

        sub = "%s　生成于 %s" % (r.get("date", ""), r.get("generated", ""))
        if r.get("quotes_updated"):
            sub += "　|　行情更新于 %s" % r["quotes_updated"][11:16]
        self.sub.text = sub
        p = r.get("products") or {}
        iv = r.get("iv") or {}
        b = r.get("breadth") or {}
        m = r.get("money") or {}

        section(c, "净持仓（前20席位）")
        for k in ("IH", "IF", "IC", "IM"):
            if k in p:
                kv(c, k, "%d" % p[k])
        kv(c, "合计", "%d" % r.get("total", 0), vcolor=C_RED)
        d = r.get("delta")
        if d is not None:
            kv(c, "较上一交易日", "%+d" % d, vcolor=(C_RED if d > 0 else C_GREEN))

        section(c, "波动率与基差")
        civ = iv.get("call_iv")
        piv = iv.get("put_iv")
        kv(c, "认购隐波", "%.2f%%" % civ if civ is not None else "—")
        kv(c, "认沽隐波", "%.2f%%" % piv if piv is not None else "—")
        if civ is not None and piv is not None:
            kv(c, "认沽溢价", "%+.2f pct" % (piv - civ))
        kv(c, "IM 基差", "%s 点" % r.get("basis") if r.get("basis") is not None else "—")
        q = r.get("quotes") or {}
        for nm in ("上证指数", "沪深300", "中证500", "中证1000"):
            if nm in q and q[nm].get("chg") is not None:
                chg = q[nm]["chg"] * 100
                kv(c, nm, "%+.2f%%" % chg, vcolor=(C_RED if chg >= 0 else C_GREEN))

        section(c, "市场宽度")
        sh, sz = b.get("sh") or ("—",) * 3, b.get("sz") or ("—",) * 3
        kv(c, "沪市 涨/跌/平", "%s / %s / %s" % tuple(sh))
        kv(c, "深市 涨/跌/平", "%s / %s / %s" % tuple(sz))
        kv(c, "主力资金", "%.2f 亿" % m["main"] if m.get("main") is not None else "—")

        if r.get("diag_rows"):
            section(c, "性质鉴别：被动套保 vs 主动做空")
            for no, nm, val, verdict in r["diag_rows"]:
                card(c, "%s %s\n%s\n→ %s" % (no, nm, val, verdict), size=13)
            if r.get("verdict"):
                card(c, "综合定性：%s" % r["verdict"], size=13, color=C_BLUE)

        if r.get("concl"):
            section(c, "结论")
            for t in r["concl"]:
                card(c, t, size=13.5)

        if r.get("watch"):
            section(c, "次日观察位")
            card(c, r["watch"], size=13.5)

        if r.get("note"):
            section(c, "数据缺口")
            card(c, r["note"], size=12.5, color=C_MUTE)


class TrackerApp(App):
    title = "股指期货持仓追踪"

    def build(self):
        self.root_widget = Root()
        return self.root_widget

    def on_start(self):
        # 启动后台定时服务（p4a）；桌面/不支持时静默跳过
        try:
            from jnius import autoclass
            Service = autoclass("org.example.indexfuturestracker.ServiceTracker")
            Activity = autoclass("org.kivy.android.PythonActivity").mActivity
            Service.start(Activity, "")
        except Exception:
            pass

    def on_pause(self):
        return True   # 切后台不销毁，回来继续


if __name__ == "__main__":
    TrackerApp().run()
