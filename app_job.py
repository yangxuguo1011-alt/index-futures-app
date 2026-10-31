# -*- coding: utf-8 -*-
"""APP 版分析流程。

设计原则：手机上跑的就是电脑端那份 daily_job.py（同一份文件复制过来的），
抓取、计算、三条判据、结论生成全部复用，只是把数据目录重定向到手机私有存储。
好处是电脑端和手机端口径永远一致，不会两套逻辑各自漂移。

与电脑端的差异（必要的两处）：
  1. 不生成 xlsx 分析文件（手机上打不开也不好看），改为返回结构化 dict，
     落盘 json 供 UI 渲染，同时导出一份 HTML 便于分享/备份。
  2. 不维护追踪主表 xlsx（read_aux_from_track 依赖它），
     昨日基差/隐波改从 outputs 里的当日存档构造，效果等价。
"""

import os
import io
import json
import shutil
import datetime
import traceback

import daily_job as dj

PACKAGE = "org.example.indexfuturestracker"

# 收盘后数据才完整（中金所持仓排名一般在 15:30 后发布）
CLOSE_HM = (15, 30)

# 盘中轻量刷新时间窗（只拉行情/宽度/资金，不重跑持仓分析）
LIVE_WINDOW = (datetime.time(9, 0), datetime.time(15, 35))


def app_dir():
    """APP 私有目录。安卓固定为 /data/data/<pkg>/files/app；桌面测试时回退脚本目录。"""
    p = "/data/data/%s/files/app" % PACKAGE
    if os.path.isdir(p):
        return p
    return os.path.dirname(os.path.abspath(__file__))


def setup_storage(base=None):
    """把 daily_job 的全局路径全部重定向到手机存储，并初始化目录与序列底稿。"""
    base = base or app_dir()
    dj.BASE = base
    dj.DATA = os.path.join(base, "data")
    dj.OUTDIR = os.path.join(base, "outputs")
    dj.ANALYSIS_DIR = os.path.join(base, "盘面多维分析")
    dj.TRACK = os.path.join(dj.DATA, "国内股指期货持仓追踪.xlsx")
    dj.SERIES = os.path.join(dj.DATA, "em_position_series.json")
    dj.SKIP_FILE = os.path.join(dj.DATA, "skipped_dates.json")
    dj.LOG = os.path.join(base, "run.log")
    for d in (dj.DATA, dj.OUTDIR, dj.ANALYSIS_DIR):
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            pass
    # 首次安装：用打包进去的历史初始化，否则净变动/趋势/判据二都无从计算
    #   seed_series.json  净持仓序列底稿（决定 H 列净变动与历史趋势）
    #   seed_outputs/     最近若干交易日的持仓明细 + 隐波存档
    #                     （判据二需要昨日明细；prev_aux 需要昨日隐波/基差）
    seed = os.path.join(base, "seed_series.json")
    try:
        if not os.path.exists(dj.SERIES) and os.path.exists(seed):
            shutil.copy(seed, dj.SERIES)
    except Exception:
        pass
    seed_out = os.path.join(base, "seed_outputs")
    try:
        if os.path.isdir(seed_out):
            for fn in os.listdir(seed_out):
                src, dst = os.path.join(seed_out, fn), os.path.join(dj.OUTDIR, fn)
                if fn.endswith(".json") and not os.path.exists(dst):
                    shutil.copy(src, dst)
    except Exception:
        pass
    return base


def _emit(progress, msg):
    if progress:
        try:
            progress(msg)
        except Exception:
            pass


def pick_trade_date(progress=None, max_back=7):
    """定位最近一个能取到持仓数据的交易日（周末/休市自动往前找）。"""
    today = datetime.date.today()
    for i in range(max_back):
        d = today - datetime.timedelta(days=i)
        ds = d.strftime("%Y-%m-%d")
        _emit(progress, "检查 %s 是否有持仓数据…" % ds)
        try:
            if dj.fetch_positions(ds):
                return ds, (i == 0)
        except Exception:
            continue
    return None, None


def _prev_aux(prev_date):
    """用昨日隐波存档构造 prev_aux（等价于电脑端从主表读的 I/J/K/L/M）。"""
    if not prev_date:
        return None
    p = os.path.join(dj.OUTDIR, "iv_%s.json" % prev_date)
    if not os.path.exists(p):
        return None
    try:
        v = json.load(io.open(p, encoding="utf-8"))
        return (prev_date, v.get("call_iv"), v.get("put_iv"),
                v.get("basis"), v.get("chg_sh"), v.get("chg_1000"))
    except Exception:
        return None


def _old_detail(prev_date):
    if not prev_date:
        return None
    p = os.path.join(dj.OUTDIR, "position_%s.json" % prev_date)
    if not os.path.exists(p):
        return None
    try:
        return json.load(io.open(p, encoding="utf-8")).get("detail")
    except Exception:
        return None


def run(progress=None, date_str=None):
    """跑一次完整分析，返回结构化结果 dict。失败 raises。"""
    setup_storage()
    ds = date_str
    if not ds:
        _emit(progress, "定位交易日…")
        ds, is_today = pick_trade_date(progress)
        if not ds:
            raise RuntimeError("最近 %d 天都未取到持仓数据（休市或网络不通）" % 7)
    _emit(progress, "交易日 %s，抓取持仓排名…" % ds)

    raw = dj.fetch_positions(ds)
    if not raw:
        raise RuntimeError("%s 无持仓数据" % ds)
    products, detail, settles = dj.calc_nets(raw)
    total = sum(products.values())

    prev_date, prev_products = dj.update_series(ds, products)
    old_products = prev_products or {k: 0 for k in dj.PRODUCTS}
    delta = total - sum(old_products.values())

    _emit(progress, "抓取行情与指数K线…")
    mkt = dj.fetch_market(ds)
    # 同一天可能被多次运行：本次没抓到的字段不得用空值冲掉上次的有效数据
    _mp = os.path.join(dj.OUTDIR, "market_%s.json" % ds)
    if os.path.exists(_mp):
        try:
            _old = json.load(io.open(_mp, encoding="utf-8"))
            for _k, _v in (("zdfb", mkt.get("zdfb")), ("ff_all", mkt.get("ff_all")),
                           ("breadth", mkt.get("breadth"))):
                if not _v and _old.get(_k):
                    mkt[_k] = _old[_k]
        except Exception:
            pass
    try:
        json.dump(mkt, io.open(_mp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    except Exception:
        pass

    quotes = dj.build_quotes(mkt, ds)
    im_close = mkt["im"]["close"]

    _emit(progress, "反解期权隐含波动率…")
    iv = {"call_iv": None, "put_iv": None, "basis": None,
          "chg_sh": quotes["上证指数"]["chg"], "chg_1000": quotes["中证1000"]["chg"]}
    if im_close:
        iv["basis"] = round(quotes["中证1000"]["close"] - im_close)
        try:
            iv = dj.fetch_iv(quotes["中证1000"]["close"], im_close)
            iv["basis"] = round(quotes["中证1000"]["close"] - im_close)
            iv["chg_sh"] = quotes["上证指数"]["chg"]
            iv["chg_1000"] = quotes["中证1000"]["chg"]
        except Exception as e:
            _emit(progress, "隐波未取到（%s），该字段留空" % type(e).__name__)
    try:
        json.dump(iv, io.open(os.path.join(dj.OUTDIR, "iv_%s.json" % ds), "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    except Exception:
        pass
    json.dump({"date": ds, "products": products, "total": total, "detail": detail},
              io.open(os.path.join(dj.OUTDIR, "position_%s.json" % ds), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # ---- 组装 build_conclusion 需要的派生量（与电脑端 build_analysis 内同源）----
    fenbu = (mkt.get("zdfb") or {}).get("fenbu") or []
    up = sum(v for x in fenbu for k, v in x.items()
             if not str(k).startswith("-") and str(k) != "0")
    down = sum(v for x in fenbu for k, v in x.items() if str(k).startswith("-"))
    flat = sum(v for x in fenbu for k, v in x.items() if str(k) == "0")
    ul = (mkt.get("breadth") or {}).get("ulist") or []
    sh_up, sh_dn, sh_flat = (ul[0]["f104"], ul[0]["f105"], ul[0]["f106"]) if ul else ("—", "—", "—")
    sz_up, sz_dn, sz_flat = (ul[1]["f104"], ul[1]["f105"], ul[1]["f106"]) if len(ul) > 1 else ("—", "—", "—")
    sb = (mkt.get("breadth") or {}).get("sina")
    if not ul and sb:
        sh_up, sh_dn, sh_flat = sb["sh"]
        sz_up, sz_dn, sz_flat = sb["sz"]
    ff = [x for x in (mkt.get("ff_all") or []) if x.startswith(ds)]
    main = small = mid = big = xbig = None
    if ff:
        main, small, mid, big, xbig = [float(x) / 1e8 for x in ff[0].split(",")[1:6]]
    basis_now = round(quotes["中证1000"]["close"] - im_close) if im_close else None

    _emit(progress, "生成结论与三条判据…")
    concl, watch, note, matrix, diag = dj.build_conclusion(
        ds, quotes, mkt, products, old_products, total, delta, iv, basis_now, im_close,
        up, down, flat, fenbu, main, xbig, big, small, _prev_aux(prev_date),
        detail=detail, old_detail=_old_detail(prev_date))

    result = {
        "date": ds,
        "generated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "products": products,
        "total": total,
        "delta": delta,
        "prev_total": sum(old_products.values()) if old_products else None,
        "prev_date": prev_date,
        "quotes": {k: {"close": v.get("close"), "chg": v.get("chg")}
                   for k, v in quotes.items()},
        "im_close": im_close,
        "iv": iv,
        "basis": basis_now,
        "breadth": {"sh": (sh_up, sh_dn, sh_flat), "sz": (sz_up, sz_dn, sz_flat),
                    "up": up or None, "down": down or None, "flat": flat or None},
        "money": {"main": main, "xbig": xbig, "big": big, "small": small},
        "concl": concl if isinstance(concl, (list, tuple)) else [concl],
        "watch": watch,
        "note": note,
        "matrix": matrix,
        "diag_rows": [list(r) for r in (diag or {}).get("rows", [])],
        "verdict": (diag or {}).get("verdict"),
        "diag_lines": dj.hedge_diag_lines(diag) if diag else [],
    }

    # 落盘：latest.json 供 UI 直接读；按日期归档一份
    try:
        json.dump(result, io.open(os.path.join(dj.OUTDIR, "latest.json"), "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
        json.dump(result, io.open(os.path.join(dj.OUTDIR, "analysis_%s.json" % ds), "w",
                                  encoding="utf-8"), ensure_ascii=False, indent=1)
        html = render_html(result)
        io.open(os.path.join(dj.ANALYSIS_DIR, "%s盘面多维分析.html" % ds), "w",
                encoding="utf-8").write(html)
    except Exception:
        pass
    _emit(progress, "完成")
    return result


def load_cached():
    """读取最近一次分析结果（UI 启动时先显示缓存，再决定是否刷新）。"""
    setup_storage()
    p = os.path.join(dj.OUTDIR, "latest.json")
    if not os.path.exists(p):
        return None
    try:
        return json.load(io.open(p, encoding="utf-8"))
    except Exception:
        return None


def _prev_weekday(d):
    """d 的上一个工作日（粗略跳过周末，不含节假日判断）。"""
    d = d - datetime.timedelta(days=1)
    while d.weekday() >= 5:
        d -= datetime.timedelta(days=1)
    return d


def refresh_quotes(progress=None):
    """盘中轻量实时刷新：只更新行情/市场宽度/主力资金/隐波基差，
    不重跑持仓分析与结论（持仓排名盘中本来也不更新）。
    非交易时段、周末或数据不可用时返回 None。成功返回合并后的完整 dict
    并同步更新 latest.json（不动 generated 时间，避免干扰收盘判定）。"""
    setup_storage()
    now = datetime.datetime.now()
    if now.weekday() >= 5:
        return None
    if not (LIVE_WINDOW[0] <= now.time() <= LIVE_WINDOW[1]):
        return None
    ds = now.strftime("%Y-%m-%d")
    _emit(progress, "实时刷新行情…")
    mkt = dj.fetch_market(ds)
    im_close = (mkt.get("im") or {}).get("close")
    if not im_close:
        return None
    quotes = dj.build_quotes(mkt, ds)

    iv = {"call_iv": None, "put_iv": None, "basis": None,
          "chg_sh": quotes["上证指数"]["chg"], "chg_1000": quotes["中证1000"]["chg"]}
    try:
        iv = dj.fetch_iv(quotes["中证1000"]["close"], im_close)
        iv["basis"] = round(quotes["中证1000"]["close"] - im_close)
        iv["chg_sh"] = quotes["上证指数"]["chg"]
        iv["chg_1000"] = quotes["中证1000"]["chg"]
    except Exception:
        pass
    basis_now = round(quotes["中证1000"]["close"] - im_close)

    fenbu = (mkt.get("zdfb") or {}).get("fenbu") or []
    up = sum(v for x in fenbu for k, v in x.items()
             if not str(k).startswith("-") and str(k) != "0")
    down = sum(v for x in fenbu for k, v in x.items() if str(k).startswith("-"))
    flat = sum(v for x in fenbu for k, v in x.items() if str(k) == "0")
    ul = (mkt.get("breadth") or {}).get("ulist") or []
    sh_up, sh_dn, sh_flat = (ul[0]["f104"], ul[0]["f105"], ul[0]["f106"]) if ul else ("—", "—", "—")
    sz_up, sz_dn, sz_flat = (ul[1]["f104"], ul[1]["f105"], ul[1]["f106"]) if len(ul) > 1 else ("—", "—", "—")
    sb = (mkt.get("breadth") or {}).get("sina")
    if not ul and sb:
        sh_up, sh_dn, sh_flat = sb["sh"]
        sz_up, sz_dn, sz_flat = sb["sz"]
    ff = [x for x in (mkt.get("ff_all") or []) if x.startswith(ds)]
    main = small = mid = big = xbig = None
    if ff:
        main, small, mid, big, xbig = [float(x) / 1e8 for x in ff[0].split(",")[1:6]]

    cached = load_cached()
    if not cached:
        return None
    cached["quotes"] = {k: {"close": v.get("close"), "chg": v.get("chg")}
                        for k, v in quotes.items()}
    cached["im_close"] = im_close
    cached["iv"] = iv
    cached["basis"] = basis_now
    cached["breadth"] = {"sh": (sh_up, sh_dn, sh_flat), "sz": (sz_up, sz_dn, sz_flat),
                         "up": up or None, "down": down or None, "flat": flat or None}
    cached["money"] = {"main": main, "xbig": xbig, "big": big, "small": small}
    cached["quotes_updated"] = now.strftime("%Y-%m-%d %H:%M:%S")
    try:
        json.dump(cached, io.open(os.path.join(dj.OUTDIR, "latest.json"), "w",
                                  encoding="utf-8"), ensure_ascii=False, indent=1)
    except Exception:
        pass
    return cached


def should_refresh():
    """是否需要完整刷新：当天还没跑过，或跑的时候还没到收盘（数据不完整）。"""
    cached = load_cached()
    if not cached:
        return True
    now = datetime.datetime.now()
    today = now.date()
    if cached.get("date") != today.strftime("%Y-%m-%d"):
        # 周末/节假日不必折腾
        if today.weekday() >= 5:
            return False
        # 收盘后：该抓当天（或最近交易日）的了
        if (now.hour, now.minute) >= CLOSE_HM:
            return True
        # 交易时段：缓存已是上一交易日就够新，不必反复重跑完整分析
        return cached.get("date") != _prev_weekday(today).strftime("%Y-%m-%d")
    try:
        gen = datetime.datetime.strptime(cached.get("generated", ""), "%Y-%m-%d %H:%M:%S")
    except Exception:
        return True
    closed = (now.hour, now.minute) >= CLOSE_HM
    was_closed = (gen.hour, gen.minute) >= CLOSE_HM
    return closed and not was_closed


# ---------------- HTML 导出（便于分享/备份） ----------------

def render_html(r):
    def esc(s):
        return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    p = r.get("products") or {}
    iv = r.get("iv") or {}
    b = r.get("breadth") or {}
    m = r.get("money") or {}
    rows = []
    for k in ("IH", "IF", "IC", "IM"):
        if k in p:
            rows.append("<tr><td>%s</td><td class='n'>%d</td></tr>" % (k, p[k]))
    diag = "".join("<li>%s</li>" % esc(x) for x in (r.get("diag_lines") or []))
    concl = "".join("<p>%s</p>" % esc(x) for x in (r.get("concl") or []))
    money = "—" if m.get("main") is None else "%.2f 亿" % m["main"]
    return u"""<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>%(date)s 盘面多维分析</title>
<style>
body{font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif;
margin:0;padding:12px;background:#f5f5f7;color:#1a1a1a;line-height:1.6}
h1{font-size:19px;margin:0 0 4px}h2{font-size:16px;margin:18px 0 8px;padding-left:8px;border-left:4px solid #3b7ddd}
.meta{color:#888;font-size:12px;margin-bottom:12px}
table{width:100%%;border-collapse:collapse;background:#fff;border-radius:8px;overflow:hidden}
td{padding:8px 10px;border-bottom:1px solid #eee;font-size:14px}
.n{text-align:right;font-variant-numeric:tabular-nums}
.card{background:#fff;border-radius:8px;padding:10px 12px;margin-bottom:10px}
.kv{display:flex;justify-content:space-between;font-size:14px;padding:3px 0}
p{font-size:14px;margin:8px 0}li{font-size:14px;margin:6px 0}
.red{color:#d93025}.green{color:#1a7f37}
</style></head><body>
<h1>%(date)s 盘面多维分析</h1>
<div class="meta">生成时间 %(gen)s　|　合约 IM%(contract)s</div>
<h2>净持仓（前20席位）</h2>
<table>%(rows)s
<tr><td><b>合计</b></td><td class="n"><b>%(total)s</b></td></tr>
<tr><td>较上一交易日</td><td class="n">%(delta)s</td></tr></table>
<h2>关键指标</h2>
<div class="card">
<div class="kv"><span>认购隐波</span><b>%(civ)s</b></div>
<div class="kv"><span>认沽隐波</span><b>%(piv)s</b></div>
<div class="kv"><span>基差</span><b>%(basis)s 点</b></div>
<div class="kv"><span>主力资金</span><b>%(money)s</b></div>
<div class="kv"><span>沪市涨/跌/平</span><b>%(sh)s</b></div>
<div class="kv"><span>深市涨/跌/平</span><b>%(sz)s</b></div>
</div>
<h2>性质鉴别（被动套保 vs 主动做空）</h2>
<ul>%(diag)s</ul>
<div class="card"><b>综合定性：</b>%(verdict)s</div>
<h2>结论</h2>%(concl)s
<h2>次日观察位</h2><p>%(watch)s</p>
%(note)s
</body></html>""" % {
        "date": r.get("date"), "gen": r.get("generated"), "contract": dj.CONTRACT,
        "rows": "".join(rows), "total": r.get("total"), "delta": r.get("delta"),
        "civ": ("%.2f%%" % iv["call_iv"]) if iv.get("call_iv") is not None else "—",
        "piv": ("%.2f%%" % iv["put_iv"]) if iv.get("put_iv") is not None else "—",
        "basis": r.get("basis") if r.get("basis") is not None else "—",
        "money": money,
        "sh": "%s / %s / %s" % b.get("sh", ("—", "—", "—")),
        "sz": "%s / %s / %s" % b.get("sz", ("—", "—", "—")),
        "diag": diag or "<li>本次未取到足够数据</li>",
        "verdict": esc(r.get("verdict") or "—"),
        "concl": concl or "<p>—</p>",
        "watch": esc(r.get("watch") or "—"),
        "note": ("<h2>数据缺口</h2><p>%s</p>" % esc(r["note"])) if r.get("note") else "",
    }
