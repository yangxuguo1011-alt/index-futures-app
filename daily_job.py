# -*- coding: utf-8 -*-
"""
国内股指期货持仓追踪 · 独立运行版
====================================
每个交易日收盘后运行（建议任务计划 16:30 触发，也可手动运行）。

功能：
  1. 抓取当日 IH/IF/IC/IM 四品种全部挂牌合约的前20席位持仓（东方财富）
  2. 抓取六指数行情（腾讯）、IM 主力合约收盘（新浪）、涨跌家数与主力资金（东财）
  3. 抓取中证1000股指期权链（新浪），Black-76 本地反解虚一档认购/认沽隐波
  4. 更新 data/国内股指期货持仓追踪.xlsx（12列，直接改 sheet1.xml，不碰格式）
  5. 生成 盘面多维分析/YYYY-MM-DD盘面多维分析.xlsx（四表结构）
  6. 为生成的 xlsx 打 AIGC 标识（docProps/custom.xml）

数据源全部为公开接口，不需要通达信、不需要任何账号。
依赖：Python 3.8+，requests（pip install requests）

休市日自动跳过：若当日持仓接口全为空，直接退出，不生成任何文件。
"""

import json
import math
import os
import re
import shutil
import sys
import time
import zipfile
import datetime

try:
    import requests
except ImportError:
    print("缺少 requests 库。")
    print("推荐直接双击 run_daily.bat（会自动准备可用的 Python 环境）；")
    print("或手动执行: pip install requests openpyxl")
    sys.exit(1)

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

# ==================== 配置 ====================
BASE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(BASE, "data")
TRACK = os.path.join(DATA, "国内股指期货持仓追踪.xlsx")
SERIES = os.path.join(DATA, "em_position_series.json")
OUTDIR = os.path.join(BASE, "outputs")
ANALYSIS_DIR = os.path.join(BASE, "盘面多维分析")
LOG = os.path.join(BASE, "run.log")

# 基差/隐波所用合约月份：11月合约 2611 于 2026-11-20（11月第三个周五）到期，
# 到期后的第一个交易日请把 CONTRACT 改为 "2612"（12月同理）。
CONTRACT = "2611"

PRODUCTS = ["IH", "IF", "IC", "IM"]
MONTHS = ["2611", "2612", "2703", "2706", "2709", "2712"]  # 自动探测，缺的返回空
WEEKDAY_CN = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
R = 0.015  # 期权反解无风险利率

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"}


def cn_date(d):
    return "%d年%d月%d日" % (d.year, d.month, d.day)


def cn(dstr):
    y, m, dd = dstr.split("-")
    return "%d年%d月%d日" % (int(y), int(m), int(dd))


def weekday_cn_from_date_cn(s):
    """'2026年10月23日' -> '星期五'（长格式，与追踪表手工填写一致）；解析失败返回 None。"""
    m = re.match(r"(\d{4})年(\d{1,2})月(\d{1,2})日", str(s).strip())
    if not m:
        return None
    try:
        d = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None
    return "星期" + "一二三四五六日"[d.weekday()]


def log(msg):
    line = "[%s] %s" % (datetime.datetime.now().strftime("%H:%M:%S"), msg)
    print(line)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def ensure_log_bom():
    """日志带 UTF-8 BOM，记事本双击打开不会乱码。"""
    try:
        if not os.path.exists(LOG) or os.path.getsize(LOG) == 0:
            with open(LOG, "w", encoding="utf-8-sig"):
                pass
    except Exception:
        pass


# ==================== 1. 持仓（东方财富） ====================
def fetch_positions(date_str):
    """返回 {code: rows}，code 如 IM2611。行内 TYPE=4 前20多头合计 / TYPE=7 前20空头合计。"""
    base = ("https://datacenter-web.eastmoney.com/api/data/v1/get?"
            "reportName=RPT_FUTU_DAILYPOSITION&columns=ALL&pageSize=500&pageNumber=1&filter=")
    data = {}
    n_all = len(PRODUCTS) * len(MONTHS)
    n_done = 0
    for prod in PRODUCTS:
        for m in MONTHS:
            code = prod + m
            flt = '(SECURITY_CODE="%s")(TRADE_DATE=\'%s\')' % (code, date_str)
            url = base + flt.replace('"', '%22').replace("'", '%27').replace('=', '%3D')
            rows = []
            for i in range(3):
                try:
                    r = requests.get(url, timeout=20, headers=dict(UA, **{"Referer": "https://data.eastmoney.com/"}))
                    rows = (r.json().get("result") or {}).get("data") or []
                    break
                except Exception as e:
                    if i == 2:
                        log("  %s 持仓抓取未成功（%s），该合约本轮跳过" % (code, _net_err(e)))
                    time.sleep(2)
            if rows:
                data[code] = rows
            n_done += 1
            if n_done % 8 == 0 or n_done == n_all:
                log("  持仓抓取进度 %d/%d（已获 %d 个合约）" % (n_done, n_all, len(data)))
            time.sleep(0.25)
    return data


def calc_nets(raw):
    """raw 为 fetch_positions 的返回。返回 (products, detail, settles)。"""
    products, detail, settles = {}, {}, {}
    for v in PRODUCTS:
        products[v] = 0
    for code, rows in sorted(raw.items()):
        longs = sum(x.get("LONG_POSITION") or 0 for x in rows if x.get("TYPE") == "4")
        shorts = sum(x.get("SHORT_POSITION") or 0 for x in rows if x.get("TYPE") == "7")
        net = longs - shorts
        products[code[:2]] = products.get(code[:2], 0) + net
        detail[code] = {"long": longs, "short": shorts, "net": net}
        st = next((x.get("SETTLE_PRICE") for x in rows if x.get("SETTLE_PRICE")), None)
        if st:
            settles[code] = st
    return products, detail, settles


def update_series(date_str, products):
    """维护 {日期: {IH,IF,IC,IM}} 序列底稿，返回 (prev_date, prev_products)。"""
    series = {}
    if os.path.exists(SERIES):
        series = json.load(open(SERIES, encoding="utf-8"))
    keys = sorted(k for k in series if k < date_str)
    prev_date = keys[-1] if keys else None
    prev = series.get(prev_date) if prev_date else None
    if date_str not in series:
        series[date_str] = {v: products.get(v, 0) for v in PRODUCTS}
        json.dump(series, open(SERIES, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        log("序列底稿已追加 %s" % date_str)
    elif series[date_str] != {v: products.get(v, 0) for v in PRODUCTS}:
        series[date_str] = {v: products.get(v, 0) for v in PRODUCTS}
        json.dump(series, open(SERIES, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
        log("序列底稿 %s 已按最新抓取结果修正" % date_str)
    return prev_date, prev


# ==================== 2. 行情（腾讯 + 新浪 + 东财） ====================
IDX_LIST = [("上证指数", "sh000001"), ("深证成指", "sz399001"), ("沪深300", "sh000300"),
            ("创业板指", "sz399006"), ("科创50", "sh000688"), ("中证1000", "sh000852")]


EM_HOSTS = ["https://push2.eastmoney.com", "https://push2delay.eastmoney.com",
            "https://80.push2.eastmoney.com"]


def _net_err(e):
    """把 requests 异常压成一句短话，日志里不出现整条异常链。"""
    try:
        s = str(e).lower()
    except Exception:
        s = ""
    if isinstance(e, requests.exceptions.Timeout) or "timeout" in s or "timed out" in s:
        return "连接超时"
    if isinstance(e, requests.exceptions.ProxyError) or "proxy" in s:
        return "代理/网络不可达"
    if isinstance(e, requests.exceptions.SSLError) or "ssl" in s:
        return "SSL 校验失败"
    if "remote end closed" in s or "connection aborted" in s or "connection reset" in s:
        return "连接被远端中断"
    return "网络连接失败"


def _em_get(path, timeout=15):
    """东财 push2 系列接口，带主机轮换重试。"""
    last = None
    for i in range(9):
        h = EM_HOSTS[i % len(EM_HOSTS)]
        try:
            return requests.get(h + path, timeout=timeout, headers=UA).json()
        except Exception as e:
            last = e
            time.sleep(1.5)
    raise last


def _sina_breadth():
    """沪/深两市涨跌平家数的备用源（东财 push2 被网络拦截时使用）。

    新浪该接口不支持直接给家数，但支持按涨跌幅降序分页；用二分查找定位
    “最后一个涨幅>0”和“最后一个>=0”的位置，即可精确得到涨/平/跌只数，
    全程约 20 次请求（每页 100 只）。
    """
    S = requests.Session()
    S.headers.update(dict(UA, **{"Referer": "https://finance.sina.com.cn/"}))
    API = ("https://vip.stock.finance.sina.com.cn/quotes_service/api/json_v2.php/"
           "Market_Center.getHQNodeData")

    def total(node):
        u = ("https://vip.stock.finance.sina.com.cn/quotes_service/api/json_v2.php/"
             "Market_Center.getHQNodeStockCount?node=" + node)
        return int(S.get(u, timeout=15).text.strip().strip('"'))

    cache = {}

    def pct_at(node, idx):
        p = (idx - 1) // 100 + 1
        if (node, p) not in cache:
            u = "%s?page=%d&num=100&sort=changepercent&asc=0&node=%s&symbol=&_s_r_a=page" % (API, p, node)
            t = S.get(u, timeout=20).text.strip()
            cache[(node, p)] = json.loads(t) if t and t != "null" else []
        rows = cache[(node, p)]
        off = (idx - 1) % 100
        return float(rows[off]["changepercent"]) if off < len(rows) else None

    out = {}
    for node, key in [("sh_a", "sh"), ("sz_a", "sz")]:
        n = total(node)
        if pct_at(node, 1) is None or pct_at(node, 1) <= 0:
            out[key] = (0, n, 0)
            continue
        lo, hi = 1, n                       # 涨：最后一个 >0
        while lo < hi:
            mid = (lo + hi + 1) // 2
            v = pct_at(node, mid)
            if v is not None and v > 0:
                lo = mid
            else:
                hi = mid - 1
        up = lo
        last0 = up
        nxt = pct_at(node, up + 1) if up < n else None
        if nxt is not None and nxt >= 0:
            lo, hi = up + 1, n              # 平：之后最后一个 >=0
            while lo < hi:
                mid = (lo + hi + 1) // 2
                v = pct_at(node, mid)
                if v is not None and v >= 0:
                    lo = mid
                else:
                    hi = mid - 1
            last0 = lo
        out[key] = (up, n - last0, last0 - up)
    return {"sh": out["sh"], "sz": out["sz"]}


def fetch_market(date_str):
    S = requests.Session()
    S.headers.update(dict(UA, **{"Referer": "https://gu.qq.com/", "Accept": "*/*"}))
    hist = {}
    for nm, code in IDX_LIST:
        for i in range(3):
            try:
                j = S.get("https://web.ifzq.gtimg.cn/appstock/app/fqkline/get?param=%s,day,,,12,qfq" % code,
                          timeout=20).json()
                node = j["data"][code]
                kl = node.get("qfqday") or node.get("day")
                hist[nm] = [{"date": x[0], "open": float(x[1]), "close": float(x[2]),
                             "high": float(x[3]), "low": float(x[4])} for x in kl]
                break
            except Exception as e:
                if i == 2:
                    log("  %s K线未取到（%s），该指数行情将留空" % (nm, _net_err(e)))
                time.sleep(2)
        time.sleep(0.2)

    im = {"hist": [], "close": None, "date": None}
    S2 = requests.Session()
    S2.headers.update(dict(UA, **{"Referer": "https://finance.sina.com.cn/"}))
    sym = "IM" + CONTRACT
    try:
        txt = S2.get("https://stock2.finance.sina.com.cn/futures/api/jsonp.php/t/"
                     "InnerFuturesNewService.getDailyKLine?symbol=" + sym, timeout=25).text
        s0, s1 = txt.index("("), txt.rindex(")")
        arr = json.loads(txt[s0 + 1:s1])
        im["hist"] = [{"date": a["d"], "close": float(a["c"]), "open": float(a["o"]),
                       "high": float(a["h"]), "low": float(a["l"])} for a in arr[-12:]]
        im["close"] = im["hist"][-1]["close"]
        im["date"] = im["hist"][-1]["date"]
    except Exception as e:
        log("  %s 日线未取到（%s），期货收盘/基差将留空" % (sym, _net_err(e)))

    E = requests.Session()
    E.headers.update(dict(UA, **{"Referer": "https://quote.eastmoney.com/"}))
    breadth, zdfb, ff_all = {}, None, None
    try:
        j = _em_get("/api/qt/ulist.np/get?secids=1.000001,0.399001"
                    "&fields=f104,f105,f106,f124&fltt=2")
        breadth["ulist"] = j.get("data", {}).get("diff")
    except Exception as e:
        log("  涨跌家数（东财）未取到（%s），尝试新浪备用源…" % _net_err(e))
        try:
            breadth["sina"] = _sina_breadth()
            log("  已改用新浪备用源统计沪深家数：沪 %s / 深 %s" % (breadth["sina"]["sh"], breadth["sina"]["sz"]))
        except Exception as e2:
            log("  新浪备用源也未取到（%s），涨跌家数将留「—」" % _net_err(e2))
    for h in ["https://push2ex.eastmoney.com", "https://push2.eastmoney.com"]:
        try:
            j = requests.get(h + "/getTopicZDFenBu?ut=7eea3edcaed734bea9cbfc24409ed989&dpt=wz.ztzt",
                             timeout=15, headers=UA).json()
            if j.get("data"):
                zdfb = j["data"]
                break
        except Exception:
            pass
    try:
        j = _em_get("/api/qt/stock/fflow/kline/get?lmt=6&klt=101"
                    "&secid=1.000001&secid2=0.399001&fields1=f1,f2,f3,f7"
                    "&fields2=f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61,f62,f63,f64,f65",
                    timeout=20)
        ff_all = j.get("data", {}).get("klines")
    except Exception as e:
        log("  主力资金未取到（%s），表格将留「—」，不影响主流程" % _net_err(e))

    return {"date": date_str, "hist": hist, "im": im, "breadth": breadth,
            "zdfb": zdfb, "ff_all": ff_all}


def build_quotes(mkt, date_str):
    """由日K重建六指数当日快照（昨收/开/高/低/收/缺口/涨跌幅/收盘位置）。"""
    quotes = {}
    for nm, _ in IDX_LIST:
        h = mkt["hist"][nm]
        td = next((x for x in h if x["date"] == date_str), None)
        if td is None:
            raise RuntimeError("%s 未取到 %s 日K" % (nm, date_str))
        i = h.index(td)
        pd_ = h[i - 1] if i > 0 else None
        if pd_ is None:
            raise RuntimeError("%s 缺前一日K线" % nm)
        q = {"prev": pd_["close"], "open": td["open"], "high": td["high"],
             "low": td["low"], "close": td["close"], "date": td["date"]}
        q["gap"] = q["open"] / q["prev"] - 1
        q["chg"] = q["close"] / q["prev"] - 1
        q["pin"] = (q["close"] - q["low"]) / (q["high"] - q["low"]) if q["high"] > q["low"] else 0.5
        quotes[nm] = q
    return quotes


# ==================== 3. 隐波（新浪期权链 + Black-76 本地反解） ====================
def ncdf(x):
    return 0.5 * (1 + math.erf(x / math.sqrt(2)))


def b76(F, K, T, sig, cp):
    if sig <= 0:
        return max(0.0, (F - K) if cp == "C" else (K - F))
    d1 = (math.log(F / K) + 0.5 * sig * sig * T) / (sig * math.sqrt(T))
    d2 = d1 - sig * math.sqrt(T)
    if cp == "C":
        return math.exp(-R * T) * (F * ncdf(d1) - K * ncdf(d2))
    return math.exp(-R * T) * (K * ncdf(-d2) - F * ncdf(-d1))


def implied(price, F, K, T, cp):
    lo, hi = 1e-4, 3.0
    for _ in range(200):
        mid = (lo + hi) / 2
        if b76(F, K, T, mid, cp) > price:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2


def _expire_from_contract(contract):
    """从合约月份（如 2611）推算到期日：到期月的第三个周五。"""
    y, m = 2000 + int(contract[:2]), int(contract[2:])
    first = datetime.date(y, m, 1)
    first_fri = first + datetime.timedelta(days=(4 - first.weekday()) % 7)
    return first_fri + datetime.timedelta(days=14)


def fetch_iv(spot, fut):
    """虚一档认购/认沽隐波（Black-76 本地反解，口径校准至同花顺期货通/博易云）。

    价格来源优先级：
      1) 中金所官方日行情收盘价（权威，与终端软件同源）；
      2) 回退新浪期权链 —— 注意取 row[2]=最新价，row[3] 是卖一价，曾误用导致偏差。

    口径校准（2026-10-29 用同花顺/博易云实显值回测 10-26/27/29 三天确定）：
      - 定价标的用 现货指数收盘价（不是期货收盘价）；
      - T = (距到期日历天数 + 1) / 365；
      - 期权价用收盘价。
      该口径与同花顺误差约 ±1 个百分点内（软件内部对期权价做了最后15分钟
      VWAP/曲面平滑，公开数据无法完全复现，此为可编程接口的极限精度）。
      虚一档 = 现价上/下方第一档行权价。不能用平值两腿直接反解——
      F 由平值价差反推时 put-call parity 决定平值认购/认沽隐波必然相等（历史 bug）。
    """
    chain = {"C": {}, "P": {}}
    F = float(fut) if fut else 0.0
    expire = _expire_from_contract(CONTRACT)
    days = (expire - datetime.date.today()).days
    T = (max(days, 1) + 1) / 365.0
    source = ""

    # ---- 1) 中金所官方日行情（收盘后发布；盘中/未发布时 404，走回退）----
    try:
        day = datetime.datetime.now().strftime("%Y%m%d")
        u = "http://www.cffex.com.cn/sj/hqsj/rtj/%s/%s/index.xml" % (day[:6], day[6:])
        r = requests.get(u, headers=UA, timeout=25)
        txt = r.content.decode("utf-8", errors="replace")

        def _fld(x, n):
            m = re.search(r"<%s>(.*?)</%s>" % (n, n), x)
            return m.group(1).strip() if m else ""

        for blk in re.findall(r"<dailydata>(.*?)</dailydata>", txt, re.S):
            iid = _fld(blk, "instrumentid")
            if iid == "IM" + CONTRACT:
                fc = _fld(blk, "closeprice")
                if fc:
                    F = float(fc)
            m2 = re.match(r"^MO%s-([CP])-(\d+)$" % CONTRACT, iid)
            if m2:
                c = _fld(blk, "closeprice")
                if c:
                    chain[m2.group(1)][int(m2.group(2))] = float(c)
        if chain["C"] and chain["P"] and F > 0:
            source = "中金所官方收盘价 + Black-76 本地反解"
    except Exception:
        pass

    # ---- 2) 回退：新浪期权链（row[2]=最新价）----
    if not (chain["C"] and chain["P"] and F > 0):
        chain = {"C": {}, "P": {}}
        S = requests.Session()
        S.headers.update(dict(UA, **{"Referer": "https://finance.sina.com.cn/"}))
        u = ("https://stock.finance.sina.com.cn/futures/api/openapi.php/OptionService.getOptionData"
             "?type=futures&product=mo&exchange=cffex&pinzhong=MO" + CONTRACT)
        j = S.get(u, timeout=25).json()
        d = j["result"]["data"]
        for row in d["up"]:
            try:
                chain["C"][float(row[7])] = float(row[2])
            except Exception:
                pass
        for row in d["down"]:
            try:
                chain["P"][float(str(row[7]).split("P")[-1])] = float(row[2])
            except Exception:
                pass
        source = "Sina OptionService(MO%s) 最新价 + Black-76 本地反解" % CONTRACT

    ks = sorted(set(chain["C"]) & set(chain["P"]))
    if not ks:
        raise RuntimeError("期权链为空")
    # 定价与选档都以现货指数为锚（校准口径）；期货收盘仅作参考字段保留
    anchor = float(spot) if spot else F
    ATM = min(ks, key=lambda k: abs(k - anchor))
    above = [k for k in ks if k > anchor]
    below = [k for k in ks if k < anchor]
    if not above or not below:
        raise RuntimeError("期权链缺少虚一档行权价（现价 %.1f）" % anchor)
    K_c, K_p = min(above), max(below)
    call_iv = round(implied(chain["C"][K_c], anchor, K_c, T, "C") * 100, 2)
    put_iv = round(implied(chain["P"][K_p], anchor, K_p, T, "P") * 100, 2)
    rows = []
    for K in ks:
        if abs(K - ATM) <= 400:
            rows.append({"K": K, "c": chain["C"][K],
                         "c_iv": round(implied(chain["C"][K], anchor, K, T, "C") * 100, 2),
                         "p": chain["P"][K],
                         "p_iv": round(implied(chain["P"][K], anchor, K, T, "P") * 100, 2)})
    return {"date": None, "source": source + "（现货标的/剩余天数+1，校准至同花顺口径）",
            "spot": spot, "fut": fut, "atm_strike": ATM, "expire_date": expire.isoformat(),
            "days": days, "T": T, "F_imp": round(anchor, 2),
            "call_strike": K_c, "put_strike": K_p,
            "call_price": chain["C"][K_c], "put_price": chain["P"][K_p],
            "call_iv": call_iv, "put_iv": put_iv, "chain": rows}


# ==================== 4. 更新追踪表（直接改 sheet1.xml，不动格式） ====================
COLS = ["B", "C", "D", "E"]  # 四品种


def _zip_read(path):
    with zipfile.ZipFile(path) as z:
        return z.namelist(), {n: z.read(n) for n in z.namelist()}


def _zip_write(path, names, items):
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            z.writestr(n, items[n])
    os.replace(tmp, path)


def _ensure_g_formula(new_row, row):
    """自愈 G 列合计公式（新布局）：补写 =IF(COUNT(Cn:Fn)=0,"",Cn+Dn+En+Fn)。
    兼容普通公式 <f> 与共享公式 <f t="shared" si="n"/>（WPS 保存后会转成共享格式，
    旧正则只认 <f>，误判无公式又插入一个 G 格 -> 一行出现两个 G54 单元格，
    Excel/WPS 显示为空 —— 2026-10-29 的教训，勿回退）。"""
    cells = re.findall(r'<c r="G%d"(?:[^>]*>.*?</c>|[^>]*/>)' % row, new_row)
    if any("<f" in c for c in cells):
        return new_row  # 公式已在（普通或共享）
    formula = 'IF(COUNT(C%d:F%d)=0,"",C%d+D%d+E%d+F%d)' % (row, row, row, row, row, row)
    # 情形1：G 格存在但为空样式格 -> 保留样式替换为公式格
    pat_empty = r'<c r="G%d"([^>/]*)/>' % row
    if re.search(pat_empty, new_row):
        return re.sub(pat_empty,
                      lambda m: '<c r="G%d"%s><f>%s</f></c>' % (row, m.group(1), formula),
                      new_row, count=1)
    # 情形2：G 格不存在 -> 插到 F 之后（或 H 之前，或行尾）
    frag = '<c r="G%d"><f>%s</f></c>' % (row, formula)
    pos = None
    for mm2 in re.finditer(r'<c r="([A-Z]+)%d"' % row, new_row):
        if mm2.group(1) in ("H", "I", "J", "K", "L", "M"):
            pos = mm2.start()
            break
    if pos is None:
        return new_row.replace("</row>", frag + "</row>")
    return new_row[:pos] + frag + new_row[pos:]


def _set_weekday_cell(new_row, row, w):
    """在行的 B 格写入星期文本（inlineStr），幂等：已有内容则不动。"""
    if re.search(r'<c r="B%d"[^>]*><(?:v|is|f)' % row, new_row):
        return new_row  # 已有内容
    # 情形1：B 格存在但为空样式格 -> 保留样式替换
    sm = re.search(r'<c r="B%d"([^>/]*)/>' % row, new_row)
    if sm:
        return re.sub(r'<c r="B%d"[^>/]*/>' % row,
                      lambda _: '<c r="B%d"%s t="inlineStr"><is><t>%s</t></is></c>' % (row, sm.group(1), w),
                      new_row, count=1)
    # 情形2：B 格不存在 -> 插在 A 格之后
    frag = '<c r="B%d" t="inlineStr"><is><t>%s</t></is></c>' % (row, w)
    am = re.search(r'<c r="A%d"[^>]*?(?:/>|</c>)' % row, new_row)
    pos = am.end() if am else 0
    return new_row[:pos] + frag + new_row[pos:]


def _is_weekday_header(v):
    """表头判定：兼容「星期」「星 期」「星期（全角空格）」等写法。"""
    return str(v or "").replace(" ", "").replace("\u3000", "") == "星期"


def ensure_weekday_fill():
    """补齐追踪表 B 列星期（新布局 B=星期）。幂等：只填空白，不动已有值。"""
    import openpyxl
    import warnings
    warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(TRACK, data_only=False)
    ws = wb.active
    if not _is_weekday_header(ws.cell(3, 2).value):
        log("警告：追踪表 B 列表头不是「星期」（实际=%r），跳过星期补齐。"
            "请确认表格已插入星期列并保存。" % ws.cell(3, 2).value)
        return
    todo = []
    for r in range(4, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None:
            continue
        v = ws.cell(r, 2).value
        if v in (None, "") or re.fullmatch(r"周[一二三四五六日]", str(v)):
            w = weekday_cn_from_date_cn(a)
            if w:
                todo.append((r, w))
    if not todo:
        log("星期列检查：无空白需要补齐")
        return
    names, items = _zip_read(TRACK)
    xml = items["xl/worksheets/sheet1.xml"].decode("utf-8")
    changed = 0
    for row, w in todo:
        m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % row, xml, re.S)
        if not m:
            continue
        old = m.group(0)
        # 先把旧短格式「周X」重写为长格式
        new_row = re.sub(r'(<c r="B%d"[^>]*><is><t>)周[一二三四五六日](</t></is></c>)' % row,
                         lambda mm: mm.group(1) + w + mm.group(2), old, count=1)
        if new_row == old:
            new_row = _set_weekday_cell(old, row, w)
        if new_row != old:
            xml = xml.replace(old, new_row, 1)
            changed += 1
    if not changed:
        return
    items["xl/worksheets/sheet1.xml"] = xml.encode("utf-8")
    _zip_write(TRACK, names, items)
    from xml.dom import minidom
    with zipfile.ZipFile(TRACK) as z:
        minidom.parseString(z.read("xl/worksheets/sheet1.xml"))
    log("星期列补齐 %d 行" % changed)


def _g_full_formula_frag(row, sattr=""):
    return '<c r="G%d"%s><f>IF(COUNT(C%d:F%d)=0,"",C%d+D%d+E%d+F%d)</f></c>' % (
        row, sattr, row, row, row, row, row, row)


def _force_full_g(new_row, row):
    """把该行 G 格规范化为「完整公式」：展开 WPS 的共享公式引用、丢弃可能为空的缓存值。
    背景（2026-10-30）：shared 引用 <f t="shared" si="n"/> 若带着 t="str" 的空 <v> 缓存，
    WPS 在手动计算模式下不重算，总净额列直接显示空白；展开成完整公式后打开必然重算。"""
    m = re.search(r'<c r="G%d"(?:[^>]*>.*?</c>|[^>]*/>)' % row, new_row)
    if not m or "<f>" in m.group(0):
        return new_row                      # 不存在或已是完整公式
    s = re.search(r's="(\d+)"', m.group(0))
    sattr = (' s="%s"' % s.group(1)) if s else ""
    return new_row.replace(m.group(0), _g_full_formula_frag(row, sattr), 1)


def _ensure_full_calc(path):
    """给 workbook.xml 的 calcPr 加 fullCalcOnLoad，打开时强制全量重算，
    防止旧的空/错缓存被直接显示。"""
    names, items = _zip_read(path)
    xml = items["xl/workbook.xml"].decode("utf-8")
    if "fullCalcOnLoad" in xml:
        return
    if "<calcPr" in xml:
        xml = re.sub(r"<calcPr ", '<calcPr fullCalcOnLoad="1" ', xml, count=1)
    else:
        xml = xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    items["xl/workbook.xml"] = xml.encode("utf-8")
    _zip_write(path, names, items)


def update_tracking(date_obj, products, total, iv):
    """定位日期行，定点写入（新布局：B=星期，C~F=四品种，G=合计公式，H=净变动，I~M=隐波/基差/涨跌幅）。
    返回 (status, row)。"""
    import openpyxl
    import warnings
    warnings.filterwarnings("ignore")
    date_cn = cn_date(date_obj)
    wb = openpyxl.load_workbook(TRACK, data_only=False)
    ws = wb.active
    if not _is_weekday_header(ws.cell(3, 2).value):
        log("错误：追踪表不是新布局（B 列应为「星期」，实际=%r）。"
            "请在表格中插入星期列并保存后重试。" % ws.cell(3, 2).value)
        return "BAD_LAYOUT", None
    row = None
    for r in range(4, ws.max_row + 1):
        if str(ws.cell(r, 1).value or "") == date_cn:
            row = r
            break
    if row is None:
        return "NO_ROW", None
    weekday = weekday_cn_from_date_cn(date_cn)
    if ws.cell(row, 3).value is not None:
        # 已填过 -> 仅做 G 列公式与 B 列星期自愈后返回
        names, items = _zip_read(TRACK)
        xml = items["xl/worksheets/sheet1.xml"].decode("utf-8")
        m0 = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % row, xml, re.S)
        if m0:
            healed = _force_full_g(
                _ensure_g_formula(_set_weekday_cell(m0.group(0), row, weekday), row), row)
            if healed != m0.group(0):
                xml = xml.replace(m0.group(0), healed, 1)
                items["xl/worksheets/sheet1.xml"] = xml.encode("utf-8")
                _zip_write(TRACK, names, items)
                _ensure_full_calc(TRACK)
                return "ALREADY+自愈", row
        return "ALREADY", row

    # 上一交易日总净额（从序列底稿取，避免依赖公式缓存）
    series = json.load(open(SERIES, encoding="utf-8"))
    dstr = date_obj.strftime("%Y-%m-%d")
    keys = sorted(k for k in series if k < dstr)
    prev = series[keys[-1]] if keys else None
    prev_total = sum(prev.values()) if prev else None
    delta = (total - prev_total) if prev_total is not None else None

    vals = {
        "C": products.get("IH"), "D": products.get("IF"),
        "E": products.get("IC"), "F": products.get("IM"),
        "I": iv.get("call_iv"), "J": iv.get("put_iv"),
        "K": iv.get("basis"), "L": iv.get("chg_sh"), "M": iv.get("chg_1000"),
    }
    if delta is not None:
        vals["H"] = delta

    # L/M 涨跌幅需按正负换样式：43=红字黑体(0.00%)、46=绿字黑体(0.00%)
    SIGN_STYLE = {"L": ("43", "46"), "M": ("43", "46")}
    INT_COLS = ("C", "D", "E", "F", "H", "K")

    names, items = _zip_read(TRACK)
    xml = items["xl/worksheets/sheet1.xml"].decode("utf-8")

    # 该行片段
    m = re.search(r'<row r="%d"[^>]*>(.*?)</row>' % row, xml, re.S)
    if not m:
        return "NO_ROW_XML", row
    row_xml = m.group(0)
    new_row = row_xml
    for col, v in vals.items():
        if v is None:
            continue
        # 已有值 -> 跳过；空样式格 -> 填值；不存在的格 -> 按列序插入
        pat_full = r'<c r="%s%d"[^>]*><v>[^<]*</v></c>' % (col, row)
        pat_empty = r'<c r="%s%d"([^>/]*)/>' % (col, row)
        if re.search(pat_full, new_row):
            continue
        if col in INT_COLS:
            sval = "%d" % round(v)
            sattr_fmt = (lambda mm: '<c r="%s%d"%s><v>%s</v></c>' % (col, row, mm.group(1), sval))
        else:
            sval = repr(round(float(v), 6))
            if col in SIGN_STYLE:
                sattr_fmt = (lambda mm: '<c r="%s%d" s="%s"><v>%s</v></c>'
                             % (col, row, SIGN_STYLE[col][0] if v > 0 else SIGN_STYLE[col][1], sval))
            else:
                sattr_fmt = (lambda mm: '<c r="%s%d"%s><v>%s</v></c>' % (col, row, mm.group(1), sval))
        if re.search(pat_empty, new_row):
            new_row = re.sub(pat_empty, sattr_fmt, new_row, count=1)
        else:
            # 单元格不存在：按列顺序插入，保持 OOXML 升序要求
            def colnum(s):
                n = 0
                for ch in s:
                    n = n * 26 + (ord(ch) - 64)
                return n
            tgt = colnum(col)
            pos = None
            for mm2 in re.finditer(r'<c r="([A-Z]+)%d"' % row, new_row):
                if colnum(mm2.group(1)) > tgt:
                    pos = mm2.start()
                    break
            frag = '<c r="%s%d"><v>%s</v></c>' % (col, row, sval)
            if pos is None:
                new_row = new_row.replace("</row>", frag + "</row>")
            else:
                new_row = new_row[:pos] + frag + new_row[pos:]
    new_row = _set_weekday_cell(new_row, row, weekday)  # B 列星期
    xml = xml.replace(row_xml, _force_full_g(_ensure_g_formula(new_row, row), row), 1)
    items["xl/worksheets/sheet1.xml"] = xml.encode("utf-8")
    _zip_write(TRACK, names, items)
    _ensure_full_calc(TRACK)

    # 校验（按实际写入的列检查：回补行隐波等拿不到历史数据会留空）
    expect = [2, 3, 4, 5, 6]
    for key, col in (("call_iv", 9), ("put_iv", 10), ("basis", 11),
                     ("chg_sh", 12), ("chg_1000", 13)):
        if iv.get(key) is not None:
            expect.append(col)
    wb2 = openpyxl.load_workbook(TRACK, data_only=False)
    ws2 = wb2.active
    ok = all(ws2.cell(row, c).value is not None for c in expect)
    if not ok:
        raise RuntimeError("追踪表写入校验失败，行 %d" % row)
    with zipfile.ZipFile(TRACK) as z:
        s = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
    from xml.dom import minidom
    minidom.parseString(s.encode("utf-8"))
    return "OK", row


# ==================== 5. 生成当日盘面多维分析 ====================
GRAY = "D9D9D9"
YELLOW = "FFF2CC"

# 判读矩阵自订口径：单日净空变动 / 合计净空 低于此比例视为噪声级扰动
NOISE_PCT = 5.0
# 单元格每行可容纳的汉字数（当日摘要 C:K 约 122 字符宽 / 判读矩阵 D:E 约 26）
CHARS_PER_LINE_9COL = 58
CHARS_PER_LINE_2COL = 13


def _row_h(text, per_line, base=19, pad=10, lo=34):
    """按字数估算合并单元格所需行高。"""
    if not text:
        return lo
    lines = -(-len(text) // per_line)
    return max(lo, lines * base + pad)


def fit_line(text, per_line):
    return _row_h(text, per_line)


# ==================== 信号性质鉴别：被动套保 vs 主动做空 ====================
# 每个交易日必须逐条过一遍，缺数据的判据如实标注「无法判定」，不得跳过、不得省略。
HEDGE_METHOD = [
    ("判据一", "被动套保基准手数",
     "被动套保的前提是「维持对冲敞口」：期货空头名义价值 ≈ 现货市值 × Beta。"
     "若现货与期货同跌，两边名义价值同步缩水，手数本应基本不变——被动盘在下跌时无需加空单。"
     "被动基准手数 = 昨日净空手数 ×(1+现货涨跌幅)÷(1+期货涨跌幅)。"
     "实际手数显著高于基准＝超额主动加空；显著低于基准＝主动减空／止盈。"),
    ("判据二", "持仓量方向（比净额更可靠）",
     "净持仓只是多空双边的大数差额，双边同增时会被掩盖（例：多+3.9万/空+4.4万，净额仅差5千）。"
     "改看单边持仓量（＝多头总量）的方向：价跌增仓＝新空进场主导；价跌减仓＝多头止损离场；"
     "价涨增仓＝新多进场主导；价涨减仓＝空头回补推动。"),
    ("判据三", "现货资金流（套保建仓必须伴随买入现货）",
     "多头套保＝买现货＋开期指空单，因此「买了股票指数就不该跌」。"
     "若主力资金净流出，说明机构在卖现货而非买，此时再加空单属双重看空，不可能是被动套保；"
     "只有主力资金净流入（至少不流出）时，净空扩大才可解释为新建套保仓位。"),
]
HEDGE_CAVEAT = (
    "口径限制：①前20席位为期货公司会员口径，不区分套保客户与投机客户，单看净持仓无法识别属性，"
    "必须以上三条交叉验证；②判据一以 IM 为样本（对应现货中证1000，基差与日线数据最完整），"
    "其余品种方向供参考；③被动套保成立的必要条件是「指数不跌或上涨＋现货资金净流入」，"
    "指数下跌时净空扩大一律先按主动加空处理，除非三条判据都指向套保。"
)


def build_hedge_diagnosis(chg_1000, im_chg, im_now, im_old, im_close,
                          detail, old_detail, main, delta, total, im_fee=200):
    """逐条判定「净空变化是被动套保还是主动做空」。

    返回 dict:
        rows    : [(编号, 名称, 实测数值, 判定语)]  顺序固定为判据一/二/三
        keys    : {"判据一":..., "判据二":..., "判据三":...}
        verdict : 综合定性一句话
    数据缺失时该条判定写「无法判定」，但条目照常输出，保证每日逐一过一遍。
    """
    rows, keys = [], {}

    # ---------- 判据一：被动套保基准手数 ----------
    no1, nm1, _b1 = HEDGE_METHOD[0]
    if im_old and im_chg is not None and im_close:
        passive_n = abs(im_old) * (1 + chg_1000) / (1 + im_chg)
        excess = abs(im_now) - passive_n
        notional = abs(excess) * im_close * im_fee / 1e8
        v1 = (f"昨 {abs(im_old):.0f} 手 → 被动基准 {passive_n:.0f} 手，"
              f"实际 {abs(im_now):.0f} 手；超额 {excess:+.0f} 手（名义约 {notional:.1f} 亿元）")
        if excess >= 300:
            k1 = "active"
            d1 = (f"超额加空 {excess:.0f} 手。现货{chg_1000 * 100:+.2f}%、期货{im_chg * 100:+.2f}%，"
                  f"敞口本已同步缩水、被动盘无需加仓，这部分只能是主动做空")
        elif excess <= -300:
            k1 = "reduce"
            d1 = (f"主动减空 {abs(excess):.0f} 手，低于维持敞口所需，"
                  f"属主动下调对冲比例／空头止盈，非被动套保")
        else:
            k1 = "passive"
            d1 = (f"与被动基准仅差 {excess:+.0f} 手（阈值±300手），"
                  f"符合「维持对冲敞口」特征，具备被动套保形态")
    else:
        k1 = "na"
        v1 = "IM 昨日净空或期货日线缺失"
        d1 = "数据缺失，本条无法判定"
    keys[no1] = k1
    rows.append((no1, nm1, v1, d1))

    # ---------- 判据二：持仓量方向 ----------
    no2, nm2, _b2 = HEDGE_METHOD[1]
    if detail and old_detail:
        oi_now = sum(x.get("long", 0) for x in detail.values())
        oi_old = sum(x.get("long", 0) for x in old_detail.values())
        sh_now = sum(x.get("short", 0) for x in detail.values())
        sh_old = sum(x.get("short", 0) for x in old_detail.values())
        doi = oi_now - oi_old
        pct = (doi / oi_old * 100) if oi_old else 0.0
        share = (abs(delta) / abs(doi) * 100) if doi else 0.0
        v2 = (f"单边持仓量 {oi_old:,}→{oi_now:,}（{doi:+,}，{pct:+.2f}%）；"
              f"多头 {(oi_now - oi_old):+,}、空头 {(sh_now - sh_old):+,}"
              + (f"；净额变动仅占增仓 {share:.1f}%" if doi else ""))
        if abs(pct) < 0.5:
            k2 = "flat"
            d2 = "持仓量基本持平，以多空换手为主，无明确增量方向"
        elif chg_1000 < 0 and doi > 0:
            k2 = "new_short"
            d2 = (f"价跌增仓＝新空进场主导（非多头离场）；净额变动只占增仓 {share:.1f}%，"
                  f"说明「净空扩大」远小于双边实际增量，主导力量是新空开仓")
        elif chg_1000 < 0:
            k2 = "long_liq"
            d2 = "价跌减仓＝多头止损离场／空头止盈，跌势或趋缓，不是新空打压"
        elif chg_1000 > 0 and doi > 0:
            k2 = "new_long"
            d2 = "价涨增仓＝新多进场主导，多头主动推进"
        else:
            k2 = "short_cover"
            d2 = "价涨减仓＝空头回补推动，持续性待观察"
    else:
        k2 = "na"
        v2 = "缺少昨日持仓明细，无法计算持仓量变化"
        d2 = "数据缺失，本条无法判定"
    keys[no2] = k2
    rows.append((no2, nm2, v2, d2))

    # ---------- 判据三：现货资金流 ----------
    no3, nm3, _b3 = HEDGE_METHOD[2]
    if main is None:
        k3 = "na"
        v3 = "主力资金接口未取到"
        d3 = "数据缺失，套保／投机属性无法交叉验证"
    else:
        v3 = f"主力资金 {main:+.2f} 亿"
        if main < 0:
            k3 = "outflow"
            if (delta or 0) < 0:
                d3 = (f"主力净流出 {-main:.0f} 亿＝现货端在减仓而非建仓，"
                      f"套保前提（买现货）不成立，配合期货加空属双重看空，排除被动套保")
            else:
                d3 = (f"主力净流出 {-main:.0f} 亿，现货抛压仍在；"
                      f"期货净空收窄多为空头止盈，不宜解读为趋势反转")
        else:
            k3 = "inflow"
            if (delta or 0) < 0:
                d3 = (f"主力净流入 {main:.0f} 亿，机构在买现货，"
                      f"此时加空单符合「买现货＋开空单」的套保建仓路径，按中性看待")
            else:
                d3 = f"主力净流入 {main:.0f} 亿，现货与期货同步转暖，偏多确认"
    keys[no3] = k3
    rows.append((no3, nm3, v3, d3))

    # ---------- 综合定性 ----------
    noise = abs(delta) / abs(total) * 100 if total else 0.0
    if (delta or 0) < 0:
        if k1 == "na" and k3 == "na":
            verdict = "关键判据缺失，套保／主动属性无法判定，勿据此下方向性结论"
        elif k1 == "active":
            if k3 == "outflow":
                verdict = "主动加空（现货减仓＋期货加空＝双重看空），不是被动套保"
            elif k3 == "inflow":
                verdict = "主动加空为主；但现货资金净流入，不排除含部分新套保建仓，按偏空且需次日验证"
            else:
                verdict = "主动加空（判据一明确超额）；现货资金数据缺失，套保属性未获交叉验证，按主动处理"
        elif k1 == "passive":
            if k3 == "outflow":
                verdict = ("手数贴合套保基准，但主力资金净流出使套保前提（买现货）不成立，"
                           "仍倾向主动加空，按套保处理的证据不足")
            elif k3 == "inflow":
                verdict = "手数贴合被动基准且现货资金净流入，符合被动套保特征，按中性看待"
            else:
                verdict = ("手数贴合被动套保基准，但现货资金缺失、套保前提（买现货）未获验证，"
                           "暂按中性看待并待补数据")
        elif k1 == "reduce":
            verdict = "IM 实际减空、合计却净空扩大，方向不一致，需查分品种结构与移仓影响"
        elif k3 == "outflow":
            verdict = "判据一缺失，但主力资金净流出、套保前提不成立，倾向主动加空"
        else:
            verdict = "判据一缺失，主力资金净流入，不排除新套保建仓，按中性看待"
    elif (delta or 0) > 0:
        if k3 == "outflow":
            verdict = "净空收窄但现货资金仍在流出，多为空头止盈，不构成趋势反转"
        elif k3 == "inflow":
            verdict = "净空收窄叠加现货资金流入，期现同步转暖，偏多确认"
        else:
            verdict = "净空收窄方向偏多，但缺现货资金验证，须看次日是否连续"
    else:
        verdict = "净空基本持平，多空均衡，等待方向确认"
    if chg_1000 < 0 and (delta or 0) < 0:
        verdict += "；指数下跌与套保建仓（买现货）方向冲突，套保解释本就不成立"
    if 0 < noise < NOISE_PCT:
        verdict += (f"；但净空变动仅占总净空 {noise:.1f}%（<{NOISE_PCT:.0f}% 噪声阈值），"
                    f"单日信号强度弱，须以后续 2~3 个交易日走向为准")
    return {"rows": rows, "keys": keys, "verdict": verdict}


def hedge_diag_lines(diag):
    """把鉴别结果格式化成当日摘要里逐行输出的文本（每条一行，严格逐一）。"""
    return ["【%s·%s】%s。→ %s" % (no, nm, v, d) for no, nm, v, d in diag["rows"]]


def build_conclusion(date_str, quotes, mkt, products, old_products, total, delta,
                     iv, basis_now, im_close, up, down, flat, fenbu,
                     main, xbig, big, small, prev_aux, detail=None, old_detail=None):
    """由当日客观数据自动生成「结论 / 次日观察位 / 判读矩阵定位」。

    返回 (concl, watch, note, matrix, diag)；note 为数据缺口说明，无缺口时为 None；
    diag 为「被动套保 vs 主动做空」三条判据的结构化结果（build_hedge_diagnosis 产出）。
    所有数值缺失时降级为定性描述，不臆造。
    """
    q1000, qsh = quotes["中证1000"], quotes["上证指数"]
    close1000, chg_1000 = q1000["close"], q1000["chg"]
    chg_sh = qsh["chg"]

    # ---- 近 N 日区间 / 均线（hist 为最近 12 根日K）----
    h1 = mkt["hist"]["中证1000"]
    hs = mkt["hist"]["上证指数"]
    cl = [x["close"] for x in h1]
    nd = len(h1)
    lo12 = min(x["low"] for x in h1)
    hi12 = max(x["high"] for x in h1)
    poss = (close1000 - lo12) / (hi12 - lo12) if hi12 > lo12 else 0.5
    ma5 = sum(cl[-5:]) / max(1, min(5, len(cl)))
    ma10 = sum(cl[-10:]) / max(1, min(10, len(cl)))
    sh_lo = min(x["low"] for x in hs)
    sh_cl = [x["close"] for x in hs]
    sh_ma10 = sum(sh_cl[-10:]) / max(1, min(10, len(sh_cl)))
    sh_ma5 = sum(sh_cl[-5:]) / max(1, min(5, len(sh_cl)))
    lvl = "低位" if poss <= 0.33 else ("高位" if poss >= 0.67 else "中位")

    # ---- 期现强弱 ----
    ih = mkt["im"]["hist"] or []
    im_prev = ih[-2]["close"] if len(ih) >= 2 else None
    im_chg = (im_close / im_prev - 1) if (im_prev and im_close) else None

    def bw(v):
        if v is None:
            return "—"
        v = int(round(v))
        return f"贴水{v}点" if v > 0 else (f"升水{-v}点" if v < 0 else "零基差")

    ph = prev_aux[1] if prev_aux else None
    pi_ = prev_aux[2] if prev_aux else None
    pj = prev_aux[3] if prev_aux else None
    db = (basis_now - pj) if (basis_now is not None and pj is not None) else None
    fut_strong = (im_chg is not None and im_chg > chg_1000)

    # ① 驱动归因
    if im_chg is None:
        rel = "IM 日线缺失，期现强弱无法比较"
    elif chg_1000 < 0:
        rel = (f"IM{CONTRACT} 跌 {-im_chg * 100:.2f}%，跌幅小于现货，期货相对抗跌" if fut_strong
               else f"IM{CONTRACT} 跌 {-im_chg * 100:.2f}%，跌幅大于现货，期货领跌")
    elif chg_1000 > 0:
        rel = (f"IM{CONTRACT} 涨 {im_chg * 100:.2f}%，涨幅大于现货，期货领涨" if fut_strong
               else f"IM{CONTRACT} 涨 {im_chg * 100:.2f}%，涨幅小于现货，期货跟涨乏力")
    else:
        rel = f"IM{CONTRACT} {im_chg * 100:+.2f}%，现货持平"

    if basis_now is None:
        btxt = "基差未取到"
    elif pj is None:
        btxt = f"基差（现货−期货）{bw(basis_now)}"
    elif abs(db) < 1:
        btxt = f"基差{bw(basis_now)}，与昨日基本持平"
    else:
        btxt = f"基差由{bw(pj)}{'收敛' if db < 0 else '走阔'}至{bw(basis_now)}（{db:+.0f}点）"

    if db is None or im_chg is None:
        interp = "期现与基差口径不完整，驱动归因以单向描述为准。"
    elif chg_1000 < 0 and db < 0 and fut_strong:
        interp = "下跌中贴水收敛、期货相对抗跌，说明抛压主要来自现货端（多头砍仓／主动抛售），期货端反而获得回补买盘。"
    elif chg_1000 < 0 and db < 0:
        interp = "下跌中贴水收敛，但期货跌幅不小于现货，更像是期货端主动平空而非现货主导。"
    elif chg_1000 < 0 and db > 0:
        interp = "下跌中贴水走阔，期货端承接了更多套保／做空压力，跌势有向期货传导的迹象。"
    elif chg_1000 > 0 and db < 0:
        interp = "上涨伴随贴水收敛，期现共振偏强，反弹质量较好。"
    elif chg_1000 > 0 and db > 0:
        interp = "上涨但贴水走阔，追涨力度不足，套保盘借反弹加空，反弹质量存疑。"
    else:
        interp = "基差持平，期现未出现明显背离。"

    # ② 持仓信号与强度
    dv = {k: products[k] - old_products.get(k, 0) for k in PRODUCTS}
    dsign = "净空收窄" if delta > 0 else ("净空扩大" if delta < 0 else "净空基本持平")
    ratio = abs(delta) / abs(total) * 100 if total else 0.0
    if ratio < NOISE_PCT:
        strength = (f"占合计净空 {ratio:.1f}%，低于 {NOISE_PCT:.0f}% 的噪声阈值，属噪声级扰动，"
                    "单日不足以定性，须以后续 2~3 个交易日净空走向为准")
    elif ratio < 10:
        strength = f"占合计净空 {ratio:.1f}%，属中等强度信号，需次日延续才能确认"
    else:
        strength = f"占合计净空 {ratio:.1f}%，幅度显著，方向性信号较强"
    top = sorted(PRODUCTS, key=lambda k: -abs(dv[k]))[:2]
    mv = "；".join(f"{k}净空{'收窄' if dv[k] > 0 else ('扩大' if dv[k] < 0 else '持平')}{abs(dv[k])}手"
                  for k in top)
    # ②b 性质鉴别：净空变化究竟是被动套保还是主动做空
    #     固定走三条判据（被动套保基准 / 持仓量方向 / 现货资金流），
    #     逐条结果由调用方在当日摘要里分行展开，缺数据也照常列出。
    im_now = products.get("IM", 0)
    im_old = old_products.get("IM")
    diag = build_hedge_diagnosis(chg_1000, im_chg, im_now, im_old, im_close,
                                 detail, old_detail, main, delta, total)
    hedge_txt = "性质鉴别（三条判据逐条见下方）：" + diag["verdict"] + "。"

    pin = q1000["pin"]
    if pin <= 0.2:
        pin_txt = f"收盘几乎贴着全天最低（日内位置 {pin * 100:.0f}%）"
    elif pin >= 0.8:
        pin_txt = f"收盘位于全天最高附近（日内位置 {pin * 100:.0f}%）"
    else:
        pin_txt = f"收于日内 {pin * 100:.0f}% 位置"

    # ③ 情绪与波动率
    def fcnt(pred):
        return sum(v for x in fenbu for k, v in x.items() if pred(int(k)))

    zd = fcnt(lambda k: k <= -9)
    d5 = fcnt(lambda k: k <= -5)
    u5 = fcnt(lambda k: k >= 5)
    up_t = up if up else "—"
    dn_t = down if down else "—"
    fl_t = flat if flat else "—"
    c_iv, p_iv = iv["call_iv"], iv["put_iv"]
    spread = p_iv - c_iv
    if ph is not None and pi_ is not None:
        iv_txt = (f"虚一档认购IV {c_iv:.2f}%（{c_iv - ph:+.2f}pp）、认沽IV {p_iv:.2f}%（{p_iv - pi_:+.2f}pp），"
                  f"认沽−认购价差 {spread:+.2f}pp（{spread - (pi_ - ph):+.2f}pp）")
    else:
        iv_txt = (f"虚一档认购IV {c_iv:.2f}%、认沽IV {p_iv:.2f}%，"
                  f"认沽−认购价差 {spread:+.2f}pp（无昨日值可比）")
    money = (f"主力资金 {main:+.2f} 亿（超大单 {xbig:+.2f}、大单 {big:+.2f}、小单 {small:+.2f}）"
             if main is not None else "主力资金接口未取到（表内显示「—」）")

    concl = (
        f"【结论】①驱动归因：中证1000 {chg_1000 * 100:+.2f}%、{rel}；{btxt}。{interp}"
        f"②持仓信号：{dsign} {abs(delta)} 手（合计 {total} 手），{strength}；其中{mv}。"
        f"中证1000 收 {close1000:.2f}，{pin_txt}，位于近{nd}日区间 {poss * 100:.0f}% 分位（{lvl}），"
        f"距MA5 {(close1000 / ma5 - 1) * 100:+.2f}%、距MA10 {(close1000 / ma10 - 1) * 100:+.2f}%。"
        f"{hedge_txt}"
        f"③情绪与波动率：全A {up_t}涨/{dn_t}跌/{fl_t}平，跌幅超9%（含跌停）{zd}家、跌超5%的{d5}家，"
        f"涨超5%仅{u5}家；{iv_txt}；{money}。"
    )

    # ---- 次日观察位 ----
    if q1000["open"] > close1000:
        res_txt = f"上方先补今日缺口 {q1000['open']:.2f}（今开），站上后才谈止跌，再看 MA5 {ma5:.2f}"
    else:
        res_txt = f"上方看 MA5 {ma5:.2f}、MA10 {ma10:.2f}"

    def _support(hist, close):
        """支撑位：取不含今日的前期低点；若今日已创新低则改给下方整数关。"""
        m = len(hist)
        prev_lo = min(x["low"] for x in hist[:-1]) if m > 1 else hist[0]["low"]
        if hist[-1]["low"] <= prev_lo:
            return math.floor(close / 100) * 100, True
        return prev_lo, False

    sup, broke = _support(h1, close1000)
    sh_sup, sh_broke = _support(hs, qsh["close"])
    if broke:
        sup_txt = (f"今日已跌破近{nd}日区间下沿，下方暂无技术支撑，"
                   f"参考整数关 {sup:.0f}（跌破则下行空间再开）")
    else:
        sup_txt = f"下方 {sup:.2f}（前{nd - 1}日低点）为第一支撑，有效跌破则下行空间打开"
    sh_sup_txt = (f"支撑 {sh_sup:.0f}（{'整数关' if sh_broke else f'前{nd - 1}日低点'}）")

    b_now = basis_now if basis_now is not None else 0
    b_in = max(10, int(b_now) - 10)
    b_out = int(b_now) + 15
    im_now = products.get("IM", 0)
    im_old = old_products.get("IM")
    im_cmp = f"今 {im_now}、昨 {im_old}" if im_old is not None else f"今 {im_now}"
    dn_word = "收窄" if delta > 0 else ("扩大" if delta < 0 else "持平")
    cont_word = "继续收窄" if delta > 0 else ("停止扩大并转收窄" if delta < 0 else "出现明确方向")
    if chg_1000 < 0:
        trigger = (f"站回 {q1000['open']:.0f} 且 IM 净空{cont_word}、基差进 {b_in} 点内＝止跌确认；"
                   f"跌破 {sup:.0f} 且隐波价差走阔＝下跌中继")
    else:
        trigger = (f"站稳 MA5 {ma5:.0f} 且 IM 净空{cont_word}＝反弹延续；"
                   f"跌破 {sup:.0f}＝反弹结束、重回弱势")

    watch = (
        f"【次日观察位】①中证1000：{sup_txt}；"
        f"{res_txt}。②上证指数：收 {qsh['close']:.2f}，{sh_sup_txt}、"
        f"压力 {sh_ma10:.2f}（MA10）、短线 {sh_ma5:.2f}（MA5）。"
        f"③基差：现{bw(basis_now)}，收敛至 {b_in} 点内＝空头回补延续（偏多）；"
        f"走阔至 {b_out} 点以上＝套保压力回归（偏空）。"
        f"④持仓：IM 净空 {im_now} 手（{im_cmp}），看能否连续第二日{dn_word}；"
        f"若明日净空反向变动 3000 手以上＝今日信号被证伪。"
        f"⑤隐波：认沽−认购价差 {spread:+.2f}pp，回落至 2pp 内＝避险需求消退；"
        f"走阔至 5pp 以上＝尾部风险抬升、回避追多。"
        f"⑥综合触发：{trigger}。"
    )

    # ---- 判读矩阵定位 ----
    net_dir = "收窄" if delta > 0 else ("扩大" if delta < 0 else "持平")
    up_dir = "指数上涨" if chg_1000 > 0 else ("指数下跌" if chg_1000 < 0 else "指数持平")
    grid = {
        ("收窄", "指数上涨"): ("D3", "低位：空头回补确认反弹，偏多修复；高位：获利对冲了结，冲高动能减弱"),
        ("收窄", "指数下跌"): ("E3", "低位：投机空头止盈，偏乐观但需看幅度；高位：正常回调，趋势未坏"),
        ("扩大", "指数上涨"): ("D4", "低位：对冲压制反弹，反弹质量差；高位：分歧加大，谨防冲高回落"),
        ("扩大", "指数下跌"): ("E4", "低位：不属被动套保（跌时敞口同步缩水，套保盘无需加空单），倾向主动加空，须主力资金验证；高位：主动看空，警惕加速下行"),
    }
    cell, txt = grid.get((net_dir, up_dir), ("E5", "多空均衡，等待方向确认"))
    hit = next((s for s in txt.split("；") if s.startswith(lvl)), txt)
    matrix = (
        f"今日定位：【净空{net_dir} {abs(delta)}手】×【{up_dir}，上证{chg_sh * 100:+.2f}%、"
        f"中证1000{chg_1000 * 100:+.2f}%】×【{lvl}：中证1000 收 {close1000:.2f}，"
        f"近{nd}日区间 {poss * 100:.0f}% 分位、距MA10 {(close1000 / ma10 - 1) * 100:+.2f}%】"
        f"→ 落 {cell}「{hit}」。{strength}。完整结论见「当日摘要·五」。"
    )

    # ---- 数据缺口 ----
    gaps = []
    if main is None:
        gaps.append("主力资金接口未取到（表内显示「—」），资金判断以市场宽度与基差替代")
    if not fenbu:
        gaps.append("涨跌分布未取到，宽度数据缺失")
    if im_chg is None:
        gaps.append("IM 日线缺失，期现强弱无法比较")
    if pj is None:
        gaps.append("追踪表无前一交易日基差，基差变化无法计算")
    note = ("【数据缺口】" + "；".join(gaps) + "。可在联网正常时重跑本脚本或手工补填。"
            if gaps else None)

    return concl, watch, note, matrix, diag


def aigc_label(path):
    """给 xlsx 写入 AIGC 隐式标识（GB 45438-2025），zip 级操作。"""
    payload = json.dumps({"Label": "1",
                          "ContentProducer": "00119111010856365494821B355",
                          "ProduceID": "%032x" % (int(time.time() * 1000) << 32 | os.getpid()),
                          "ContentPropagator": "00119111010856365494821B355",
                          "PropagateID": "%032x" % (int(time.time() * 1000) << 32 | os.getpid())},
                         separators=(",", ":"))
    CUSTOM = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
              '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" '
              'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
              '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="AIGC">'
              '<vt:lpwstr>%s</vt:lpwstr></property></Properties>') % payload.replace("&", "&amp;").replace('"', "&quot;")
    names, items = _zip_read(path)
    items["docProps/custom.xml"] = CUSTOM.encode("utf-8")
    if "docProps/custom.xml" not in names:
        names = list(names) + ["docProps/custom.xml"]
    ct = items["[Content_Types].xml"].decode("utf-8")
    if "custom-properties" not in ct:
        ct = ct.replace("</Types>",
                        '<Override PartName="/docProps/custom.xml" ContentType='
                        '"application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>')
        items["[Content_Types].xml"] = ct.encode("utf-8")
    rels = items["_rels/.rels"].decode("utf-8")
    if "custom.xml" not in rels:
        rels = rels.replace("</Relationships>",
                            '<Relationship Id="rIdAIGC" Type="http://schemas.openxmlformats.org/'
                            'officeDocument/2006/relationships/custom-properties" '
                            'Target="docProps/custom.xml"/></Relationships>')
        items["_rels/.rels"] = rels.encode("utf-8")
    _zip_write(path, names, items)


def build_analysis(date_obj, quotes, mkt, products, detail, old_products, settles,
                   series, aux, iv, total, delta, out_path, old_detail=None):
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

    date_str = date_obj.strftime("%Y-%m-%d")
    wd = WEEKDAY_CN[date_obj.weekday()]
    V4 = PRODUCTS
    OFF = 2
    F12 = Font(name="黑体", size=12)
    F12B = Font(name="黑体", size=12, bold=True)
    F16B = Font(name="黑体", size=16, bold=True)
    CTR = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin")
    BD = Border(left=thin, right=thin, top=thin, bottom=thin)
    GRAYF = PatternFill("solid", fgColor=GRAY)
    YELLOWF = PatternFill("solid", fgColor=YELLOW)
    REDF = PatternFill("solid", fgColor="FCE4EC")
    GREENF = PatternFill("solid", fgColor="E8F5E9")
    NUM = '#,##0_);[Red](#,##0)'
    PCT = '0.00%'

    def put(ws, r, c, v, font=None, fmt=None, fill=None, align=None, border=True):
        cell = ws.cell(row=r, column=c + OFF, value=v)
        cell.font = font or F12
        cell.alignment = align or CTR
        if border:
            cell.border = BD
        if fmt:
            cell.number_format = fmt
        if fill:
            cell.fill = fill
        return cell

    def merge(ws, r, c1, c2):
        ws.merge_cells(start_row=r, start_column=c1 + OFF, end_row=r, end_column=c2 + OFF)

    def block(ws, r, text, ncol=9, h=34, font=None, fill=None):
        put(ws, r, 1, text, font=font or F12, align=LEFT, fill=fill)
        merge(ws, r, 1, ncol)
        ws.row_dimensions[r].height = h

    # ---- 派生量 ----
    idx_order = [n for n, _ in IDX_LIST]
    im_close = mkt["im"]["close"]
    basis_now = round(quotes["中证1000"]["close"] - im_close) if im_close else None
    fenbu = (mkt.get("zdfb") or {}).get("fenbu") or []
    up = sum(v for x in fenbu for k, v in x.items() if not str(k).startswith("-") and str(k) != "0")
    down = sum(v for x in fenbu for k, v in x.items() if str(k).startswith("-"))
    flat = sum(v for x in fenbu for k, v in x.items() if str(k) == "0")
    ul = (mkt.get("breadth") or {}).get("ulist") or []
    sh_up, sh_dn, sh_flat = (ul[0]["f104"], ul[0]["f105"], ul[0]["f106"]) if len(ul) > 0 else ("—", "—", "—")
    sz_up, sz_dn, sz_flat = (ul[1]["f104"], ul[1]["f105"], ul[1]["f106"]) if len(ul) > 1 else ("—", "—", "—")
    sb = (mkt.get("breadth") or {}).get("sina")
    if not ul and sb:                       # 东财 push2 不可用时的备用口径
        sh_up, sh_dn, sh_flat = sb["sh"]
        sz_up, sz_dn, sz_flat = sb["sz"]
    ff = [x for x in (mkt.get("ff_all") or []) if x.startswith(date_str)]
    main = small = mid = big = xbig = None
    if ff:
        main, small, mid, big, xbig = [float(x) / 1e8 for x in ff[0].split(",")[1:6]]
    chg_sh = quotes["上证指数"]["chg"]
    chg_1000 = quotes["中证1000"]["chg"]
    prev_aux = aux[-2] if len(aux) >= 2 else None  # (日期, H, I, J, K, L)

    wb = openpyxl.Workbook()

    # ============ Sheet1 当日摘要 ============
    ws = wb.active
    ws.title = "当日摘要"
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 3
    ws.column_dimensions["C"].width = 18
    for col in "DEFGHIJK":
        ws.column_dimensions[col].width = 13
    r = 1
    put(ws, r, 1, "%s（%s）盘面多维分析 · 当日摘要" % (cn_date(date_obj), wd), F16B, fill=GRAYF, align=LEFT)
    merge(ws, r, 1, 9)
    ws.row_dimensions[r].height = 30
    r += 2

    put(ws, r, 1, "一、指数行情（六指数）", F12B, align=LEFT)
    merge(ws, r, 1, 9)
    r += 1
    hdr = ["指数", "昨收", "开盘", "缺口", "最高", "最低", "收盘", "涨跌幅", "收盘位置"]
    for j, h in enumerate(hdr):
        put(ws, r, 1 + j, h, F12B, fill=GRAYF)
    r += 1
    for nm in idx_order:
        q = quotes[nm]
        put(ws, r, 1, nm, F12B)
        put(ws, r, 2, q["prev"], fmt='0.00')
        put(ws, r, 3, q["open"], fmt='0.00')
        put(ws, r, 4, q["gap"], fmt=PCT, fill=(REDF if q["gap"] > 0 else GREENF))
        put(ws, r, 5, q["high"], fmt='0.00')
        put(ws, r, 6, q["low"], fmt='0.00')
        put(ws, r, 7, q["close"], fmt='0.00', font=F12B)
        put(ws, r, 8, q["chg"], fmt=PCT, fill=(REDF if q["chg"] > 0 else GREENF))
        put(ws, r, 9, q["pin"], fmt=PCT, fill=(REDF if q["pin"] >= 0.5 else GREENF))
        r += 1
    hi_pin = max(quotes[n]["pin"] for n in idx_order)
    lo_pin = min(quotes[n]["pin"] for n in idx_order)
    tone = "全线收于日内高位" if hi_pin >= 0.8 else ("全线收于日内低位" if lo_pin <= 0.2 else "涨跌互现、强弱分化")
    block(ws, r, "注：缺口=开盘/昨收-1；收盘位置=(收盘-最低)/(最高-最低)，越大越接近收在日内最高。今日六指数%s（收盘位置%.0f%%~%.0f%%）。"
          % (tone, lo_pin * 100, hi_pin * 100), 9, 30)
    r += 2

    put(ws, r, 1, "二、市场宽度与主力资金", F12B, align=LEFT)
    merge(ws, r, 1, 9)
    r += 1
    rows_b = [
        ("全A上涨家数（东财全市场口径）", up or "—", NUM),
        ("全A下跌家数（东财全市场口径）", down or "—", NUM),
        ("全A平盘家数", flat or "—", NUM),
        ("涨跌家数比（跌/涨）", round(down / up, 2) if up else "—", '0.00'),
        ("沪深两市上涨/下跌", "%s / %s" % (sh_up + sz_up if isinstance(sh_up, int) else "—",
                                          sh_dn + sz_dn if isinstance(sh_dn, int) else "—"), None),
        ("其中：沪市涨/跌/平", "%s / %s / %s" % (sh_up, sh_dn, sh_flat), None),
        ("其中：深市涨/跌/平", "%s / %s / %s" % (sz_up, sz_dn, sz_flat), None),
        ("主力资金净流入（亿元，全A）", round(main, 2) if main is not None else "—", '#,##0.00'),
        ("超大单净流入（亿元）", round(xbig, 2) if xbig is not None else "—", '#,##0.00'),
        ("大单净流入（亿元）", round(big, 2) if big is not None else "—", '#,##0.00'),
        ("中单净流入（亿元）", round(mid, 2) if mid is not None else "—", '#,##0.00'),
        ("小单净流入（亿元）", round(small, 2) if small is not None else "—", '#,##0.00'),
    ]
    for k, v, fmt in rows_b:
        put(ws, r, 1, k, fill=GRAYF, align=LEFT)
        merge(ws, r, 1, 2)
        c = put(ws, r, 3, v, fmt=fmt)
        merge(ws, r, 3, 4)
        c.alignment = CTR
        if isinstance(v, float):
            c.fill = REDF if v > 0 else GREENF
        r += 1
    block(ws, r, "注：全A家数取东财涨跌分布全市场口径（合计%d家）；资金流为东财全A口径（主力=超大单+大单），单位亿元，正值=净流入。"
          % (up + down + flat) if up else "注：当日涨跌分布数据未取到。", 9, 28)
    r += 2

    put(ws, r, 1, "三、四品种前20席位净持仓合计（净持仓=前20多头-前20空头）", F12B, align=LEFT)
    merge(ws, r, 1, 9)
    r += 1
    put(ws, r, 1, "项目", F12B, fill=GRAYF)
    for j, h in enumerate(["IH上证50", "IF沪深300", "IC中证500", "IM中证1000", "四品种合计"]):
        put(ws, r, 2 + j, h, F12B, fill=GRAYF)
    r += 1
    put(ws, r, 1, "净持仓（%s）" % date_str[5:].replace("-", "-"), fill=GRAYF, align=LEFT)
    for j, v in enumerate([products[k] for k in V4]):
        put(ws, r, 2 + j, v, fmt=NUM, fill=(REDF if v > 0 else None))
    put(ws, r, 6, total, fmt=NUM, font=F12B)
    r += 1
    put(ws, r, 1, "较上一交易日变化", fill=GRAYF, align=LEFT)
    for j, k in enumerate(V4):
        v = products[k] - old_products.get(k, 0)
        put(ws, r, 2 + j, v, fmt=NUM, fill=(REDF if v > 0 else GREENF))
    put(ws, r, 6, delta, fmt=NUM, font=F12B, fill=(REDF if (delta or 0) > 0 else GREENF))
    r += 1
    put(ws, r, 1, "前20多头合计", fill=GRAYF, align=LEFT)
    for j, k in enumerate(V4):
        put(ws, r, 2 + j, sum(d["long"] for c, d in detail.items() if c.startswith(k)), fmt=NUM)
    r += 1
    put(ws, r, 1, "前20空头合计", fill=GRAYF, align=LEFT)
    for j, k in enumerate(V4):
        put(ws, r, 2 + j, sum(d["short"] for c, d in detail.items() if c.startswith(k)), fmt=NUM)
    r += 2

    put(ws, r, 1, "四、辅助指标：11月基差 / 沪指 / 中证1000 / 虚一档隐波（当日值 + 近期趋势）", F12B, align=LEFT)
    merge(ws, r, 1, 9)
    r += 1
    put(ws, r, 1, "指标", F12B, fill=GRAYF)
    put(ws, r, 2, "当日值", F12B, fill=GRAYF)
    put(ws, r, 3, "较前日", F12B, fill=GRAYF)
    merge(ws, r, 3, 4)
    put(ws, r, 5, "说明", F12B, fill=GRAYF)
    merge(ws, r, 5, 9)
    r += 1
    ph = prev_aux[1] if prev_aux else None
    pi_ = prev_aux[2] if prev_aux else None
    pj = prev_aux[3] if prev_aux else None
    info = [
        ("11月基差（点）", basis_now, (basis_now - pj) if (basis_now is not None and pj is not None) else None,
         "中证1000收盘%.2f − IM%s收盘%s = %s（现货减期货，正=期货贴水）"
         % (quotes["中证1000"]["close"], CONTRACT, im_close, basis_now)),
        ("沪指涨跌幅", chg_sh, None, "开盘%.2f%%、最高%.2f%%、最低%.2f%%、收盘%.2f%%"
         % (quotes["上证指数"]["gap"] * 100, quotes["上证指数"]["high"] / quotes["上证指数"]["prev"] * 100 - 100,
            quotes["上证指数"]["low"] / quotes["上证指数"]["prev"] * 100 - 100, chg_sh * 100)),
        ("中证1000涨跌幅", chg_1000, None, "开盘%.2f%%、最高%.2f%%、最低%.2f%%、收盘%.2f%%、日内振幅%.1f%%"
         % (quotes["中证1000"]["gap"] * 100, quotes["中证1000"]["high"] / quotes["中证1000"]["prev"] * 100 - 100,
            quotes["中证1000"]["low"] / quotes["中证1000"]["prev"] * 100 - 100, chg_1000 * 100,
            (quotes["中证1000"]["high"] - quotes["中证1000"]["low"]) / quotes["中证1000"]["prev"] * 100)),
        ("11月认购隐波(%)", iv["call_iv"], (iv["call_iv"] - ph) if ph is not None else None,
         "MO%s虚一档（行权价%.0f，高于现价第一档）认购，Black-76本地反解；到期日 %s，剩 %d 天"
         % (CONTRACT, iv.get("call_strike", iv["atm_strike"]), iv.get("expire_date") or "—", iv["days"])),
        ("11月认沽隐波(%)", iv["put_iv"], (iv["put_iv"] - pi_) if pi_ is not None else None,
         "MO%s虚一档（行权价%.0f，低于现价第一档）认沽" % (CONTRACT, iv.get("put_strike", iv["atm_strike"]))),
        ("认沽-认购隐波差(百分点)", round(iv["put_iv"] - iv["call_iv"], 2),
         (round(iv["put_iv"] - iv["call_iv"], 2) - round(pi_ - ph, 2)) if (ph is not None and pi_ is not None) else None,
         "正值=认沽相对认购更贵（避险需求高）；负值=偏投机"),
    ]
    for nm, v, dv, note in info:
        put(ws, r, 1, nm, fill=GRAYF, align=LEFT)
        fmt = PCT if nm in ("沪指涨跌幅", "中证1000涨跌幅") else ('0.00' if "隐波" in nm or "隐波差" in nm else '0')
        c = put(ws, r, 2, v, fmt=fmt)
        if nm in ("沪指涨跌幅", "中证1000涨跌幅") and isinstance(v, float):
            c.fill = REDF if v > 0 else GREENF
        if dv is not None:
            dfmt = PCT if nm in ("沪指涨跌幅", "中证1000涨跌幅") else '0.00'
            put(ws, r, 3, round(dv, 4) if fmt == PCT else round(dv, 2), fmt=dfmt,
                fill=(REDF if dv > 0 else GREENF))
        else:
            put(ws, r, 3, "—")
        merge(ws, r, 3, 4)
        put(ws, r, 5, note, align=LEFT)
        merge(ws, r, 5, 9)
        r += 1
    r += 1

    put(ws, r, 1, "近期趋势（源：持仓追踪表）", F12B, fill=GRAYF, align=LEFT)
    merge(ws, r, 1, 9)
    r += 1
    for j, h in enumerate(["日期", "11月基差", "沪指涨跌", "中证1000涨跌", "认购隐波", "认沽隐波", "认沽-认购"]):
        put(ws, r, 1 + j, h, F12B, fill=GRAYF)
    r += 1
    for d, h, i_, j_, k_, l_ in aux:
        put(ws, r, 1, d)
        put(ws, r, 2, j_, fmt='0')
        put(ws, r, 3, k_, fmt=PCT, fill=(REDF if (k_ or 0) > 0 else GREENF))
        put(ws, r, 4, l_, fmt=PCT, fill=(REDF if (l_ or 0) > 0 else GREENF))
        put(ws, r, 5, h, fmt='0.00')
        put(ws, r, 6, i_, fmt='0.00')
        if h and i_:
            put(ws, r, 7, round(i_ - h, 2), fmt='0.00', fill=(REDF if i_ - h > 0 else GREENF))
        r += 1
    block(ws, r, "注：基差为现货减期货，正值=期货贴水；隐波单位为%。", 9, 26)
    r += 2

    # ---- 五、信号组合（自动客观部分 + 人工解读占位） ----
    put(ws, r, 1, "五、信号组合与当日结论", F12B, align=LEFT)
    merge(ws, r, 1, 9)
    r += 1
    dsign = "净空扩大 %d手" % abs(delta) if (delta or 0) < 0 else ("净空收窄 %d手" % delta if (delta or 0) > 0 else "基本持平")
    dsign2 = "指数上涨" if chg_1000 > 0 else ("指数下跌" if chg_1000 < 0 else "指数持平")
    matrix_line = {
        ("净空收窄", "指数上涨"): "空头回补/对冲比例下调（可靠性高）",
        ("净空收窄", "指数下跌"): "投机空头止盈，偏乐观（可靠性中高）",
        ("净空扩大", "指数上涨"): "低位=对冲压制反弹、反弹质量存疑；高位=主动看空/锁收益（可靠性高）",
        ("净空扩大", "指数下跌"): "低位=不属被动套保，倾向主动加空（须主力资金验证）；高位=主动看空，警惕加速下行",
    }.get((dsign.split(" ")[0], dsign2), "多空均衡，等待方向确认")
    concl = [
        "信号组合：【%s】×【%s，上证%+.2f%%、中证1000%+.2f%%】→ 矩阵判读：%s。位置（高/中/低位）由脚本按近12日区间与均线自动判定，见下方结论与「判读矩阵·今日定位」。"
        % (dsign, dsign2, chg_sh * 100, chg_1000 * 100, matrix_line),
        "持仓结构（客观）：%s。" % "；".join(
            "%s%+.0f（较昨%+.0f）" % (k, products[k], products[k] - old_products.get(k, 0)) for k in V4),
        "资金与宽度（客观）：全A上涨%s家、下跌%s家；主力资金%s。"
        % (up or "—", down or "—",
           ("%+.2f亿（超大单%+.2f、大单%+.2f）" % (main, xbig, big)) if main is not None else "—"),
        "波动率与基差（客观）：MO%s虚一档认购IV %.2f%%（K=%.0f）、认沽IV %.2f%%（K=%.0f）（价差%+.2f个百分点）；11月基差%s点。"
        % (CONTRACT, iv["call_iv"], iv.get("call_strike", 0), iv["put_iv"], iv.get("put_strike", 0),
           iv["put_iv"] - iv["call_iv"], basis_now if basis_now is not None else "—"),
    ]

    # ---- 自动生成「结论 / 性质鉴别三条判据 / 次日观察位 / 数据缺口」 ----
    gen_concl, gen_watch, gen_note, gen_matrix, gen_diag = build_conclusion(
        date_str, quotes, mkt, products, old_products, total, delta, iv, basis_now,
        im_close, up, down, flat, fenbu, main, xbig, big, small, prev_aux,
        detail=detail, old_detail=old_detail)
    gen_diag_lines = hedge_diag_lines(gen_diag)
    base_len = len(concl)
    concl += [gen_concl]
    concl += gen_diag_lines                       # 判据一 / 二 / 三，逐条各占一行
    concl += ["【性质鉴别·综合】" + gen_diag["verdict"] + "。"]
    concl += [gen_watch]
    if gen_note:
        concl.append(gen_note)

    # ---- 人工改写保护：旧文件里的结论/判据若已被人工改写则保留 ----
    cache_path = os.path.join(DATA, "conclusion_cache.json")
    cache = {}
    if os.path.exists(cache_path):
        try:
            cache = json.load(open(cache_path, encoding="utf-8"))
        except Exception:
            cache = {}
    manual = False
    old_lines = {"concl": "", "diag": [], "verdict": "", "watch": "", "note": ""}
    if os.path.exists(out_path):
        try:
            import openpyxl as _op
            _wb = _op.load_workbook(out_path)
            _ws = _wb["当日摘要"]
            # 按文本前缀定位，避免依赖固定行号（插入/删除行都不会错位）
            for _rr in range(50, _ws.max_row + 1):
                _v = _ws.cell(_rr, 3).value
                if not _v:
                    continue
                _v = str(_v)
                if _v.startswith("【结论】"):
                    old_lines["concl"] = _v
                elif _v.startswith("【判据"):
                    old_lines["diag"].append(_v)
                elif _v.startswith("【性质鉴别·综合】"):
                    old_lines["verdict"] = _v
                elif _v.startswith("【次日观察位】"):
                    old_lines["watch"] = _v
                elif _v.startswith("【数据缺口】"):
                    old_lines["note"] = _v
            cur = old_lines["concl"]
            last = cache.get(date_str) or {}
            if "请结合上述客观数据补充主观判读" in cur:
                manual = False                    # 旧占位文案，直接覆盖
            elif not old_lines["diag"]:
                manual = False                    # 旧版格式（无判据行），按新版重建
            elif last.get("concl") is not None:
                manual = bool(cur) and (cur != last.get("concl")
                                        or old_lines["diag"] != (last.get("diag") or []))
            else:
                manual = bool(cur)                # 无缓存且有内容：保守保留
            if manual:
                log("  结论/判据行已被人工改写，本次保留人工内容（自动生成稿见 conclusion_cache.json）")
                concl = concl[:base_len]
                if cur:
                    concl.append(cur)
                concl += [x for x in old_lines["diag"] if x]
                if old_lines["verdict"]:
                    concl.append(old_lines["verdict"])
                if old_lines["watch"]:
                    concl.append(old_lines["watch"])
                if old_lines["note"]:
                    concl.append(old_lines["note"])
                try:
                    _m = _wb["判读矩阵"].cell(7, 4).value
                    if _m:
                        gen_matrix = _m
                except Exception:
                    pass
        except Exception as e:
            log("  读取旧结论失败（按自动生成处理）: %r" % (e,))
    try:
        cache[date_str] = {"concl": gen_concl, "watch": gen_watch, "note": gen_note,
                           "diag": gen_diag_lines, "verdict": gen_diag["verdict"]}
        json.dump(cache, open(cache_path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    except Exception as e:
        log("  结论缓存写入失败（不影响主流程）: %r" % (e,))

    for t in concl:
        block(ws, r, t, 9, fit_line(t, CHARS_PER_LINE_9COL))
        r += 1

    # ============ Sheet2 持仓明细 ============
    ws2 = wb.create_sheet("持仓明细")
    ws2.column_dimensions["A"].width = 3
    ws2.column_dimensions["B"].width = 3
    ws2.column_dimensions["C"].width = 14
    ws2.column_dimensions["D"].width = 14
    for col in "EFGHI":
        ws2.column_dimensions[col].width = 15
    put(ws2, 1, 1, "%s 股指期货前20席位持仓明细（多/空为前20名会员合计，净持仓=多头-空头）" % cn_date(date_obj),
        F12B, fill=GRAYF, align=LEFT)
    merge(ws2, 1, 1, 7)
    for j, h in enumerate(["品种", "合约", "前20多头", "前20空头", "净持仓", "较昨日变化", "结算价"]):
        put(ws2, 2, 1 + j, h, F12B, fill=GRAYF)
    r2 = 3
    for v in V4:
        for m_ in MONTHS:
            code = v + m_
            if code not in detail:
                put(ws2, r2, 1, v)
                put(ws2, r2, 2, code)
                put(ws2, r2, 3, "当日无数据（合约未上市或尚未纳入前20统计）", align=LEFT)
                merge(ws2, r2, 3, 7)
                r2 += 1
                continue
            d = detail[code]
            p = (old_detail or {}).get(code)
            put(ws2, r2, 1, v)
            put(ws2, r2, 2, code)
            put(ws2, r2, 3, d["long"], fmt=NUM)
            put(ws2, r2, 4, d["short"], fmt=NUM)
            put(ws2, r2, 5, d["net"], fmt=NUM, fill=(REDF if d["net"] > 0 else None))
            if p:
                chg = d["net"] - p["net"]
                put(ws2, r2, 6, chg, fmt=NUM, fill=(REDF if chg > 0 else GREENF))
            else:
                put(ws2, r2, 6, "首日纳入", align=LEFT)
            if settles.get(code):
                put(ws2, r2, 7, settles[code], fmt='0.0')
            r2 += 1
        put(ws2, r2, 1, v + " 合计", F12B, fill=GRAYF)
        put(ws2, r2, 2, "", fill=GRAYF)
        put(ws2, r2, 3, sum(d["long"] for c, d in detail.items() if c.startswith(v)), F12B, fmt=NUM, fill=GRAYF)
        put(ws2, r2, 4, sum(d["short"] for c, d in detail.items() if c.startswith(v)), F12B, fmt=NUM, fill=GRAYF)
        put(ws2, r2, 5, products[v], F12B, fmt=NUM, fill=GRAYF)
        dv_ = products[v] - old_products.get(v, 0)
        put(ws2, r2, 6, dv_, F12B, fmt=NUM, fill=(REDF if dv_ > 0 else GREENF))
        put(ws2, r2, 7, "", fill=GRAYF)
        r2 += 1
    put(ws2, r2, 1, "四品种总净额", F12B, fill=YELLOWF)
    for c in (2, 3, 4, 7):
        put(ws2, r2, c, "", fill=YELLOWF)
    put(ws2, r2, 5, total, F12B, fmt=NUM, fill=YELLOWF)
    put(ws2, r2, 6, delta, F12B, fmt=NUM, fill=YELLOWF)
    r2 += 2
    block(ws2, r2, "注：变化为正=净空收窄/净多增加。结算价来自东财持仓接口（SETTLE_PRICE），仅作参考，不用于基差计算（基差用新浪收盘价）。", 7, 30)

    # ============ Sheet3 净持仓序列（近15个交易日） ============
    ws3 = wb.create_sheet("净持仓序列")
    ws3.column_dimensions["A"].width = 3
    ws3.column_dimensions["B"].width = 3
    ws3.column_dimensions["C"].width = 16
    for col in "DEFGHI":
        ws3.column_dimensions[col].width = 14
    put(ws3, 1, 1, "四品种前20席位净持仓合计序列（近15个交易日；完整历史累计维护于 data/em_position_series.json 与持仓追踪表）",
        F12B, fill=GRAYF, align=LEFT)
    merge(ws3, 1, 1, 7)
    for j, h in enumerate(["日期", "IH上证50", "IF沪深300", "IC中证500", "IM中证1000", "四品种合计", "净变动"]):
        put(ws3, 2, 1 + j, h, F12B, fill=GRAYF)
    r3 = 3
    keys_sorted = sorted(series.keys())
    last15 = keys_sorted[-15:]
    prev_total = None
    for dstr in last15:
        s_ = series[dstr]
        tot = sum(s_[k] for k in V4)
        hl = YELLOWF if dstr == date_str else None
        put(ws3, r3, 1, cn(dstr), fill=hl)
        for j, k in enumerate(V4):
            put(ws3, r3, 2 + j, s_[k], fmt=NUM, fill=hl)
        put(ws3, r3, 6, tot, fmt=NUM, font=F12B, fill=hl)
        if prev_total is not None:
            chg = tot - prev_total
            put(ws3, r3, 7, chg, fmt=NUM, fill=(REDF if chg > 0 else GREENF))
        else:
            put(ws3, r3, 7, "—")
        prev_total = tot
        r3 += 1
    peak = min(keys_sorted, key=lambda d_: sum(series[d_].values()))
    peakv = sum(series[peak].values())
    narrow = sum(1 for a, b in zip(keys_sorted[:-1], keys_sorted[1:])
                 if sum(series[b].values()) - sum(series[a].values()) > 0)
    block(ws3, r3 + 1, "注：净变动=当日四品种总净额−上一交易日，正值代表净空收窄。全样本（%d个交易日）净空峰值 %s（%d手），今日为 %s 手，较峰值收敛 %d 手；全样本中净空收窄%d次、扩大%d次。"
          % (len(keys_sorted), cn(peak), peakv, total, total - peakv, narrow, len(keys_sorted) - 1 - narrow), 7, 40)

    # ============ Sheet4 判读矩阵 ============
    ws4 = wb.create_sheet("判读矩阵")
    ws4.column_dimensions["A"].width = 3
    ws4.column_dimensions["B"].width = 3
    ws4.column_dimensions["C"].width = 16
    ws4.column_dimensions["D"].width = 46
    ws4.column_dimensions["E"].width = 46
    put(ws4, 1, 1, "净空变化 × 指数涨跌 × 发生位置 判读矩阵", F12B, fill=GRAYF, align=LEFT)
    merge(ws4, 1, 1, 3)
    for j, h in enumerate(["净空变化＼指数", "指数上涨", "指数下跌"]):
        put(ws4, 2, 1 + j, h, F12B, fill=GRAYF)
    matrix = [
        ("净空收窄", "低位：空头回补确认反弹，偏多修复；高位：获利对冲了结，冲高动能减弱",
         "低位：投机空头止盈，偏乐观但需看幅度；高位：正常回调，趋势未坏"),
        ("净空扩大", "低位：对冲压制反弹，反弹质量差；高位：分歧加大，谨防冲高回落",
         "低位：不属被动套保（跌时敞口同步缩水，套保盘无需加空单），倾向主动加空，须主力资金验证；高位：主动看空，警惕加速下行"),
        ("基本持平", "方向未确认，观望为主", "多空均衡，等待方向确认"),
    ]
    r4 = 3
    for row_ in matrix:
        put(ws4, r4, 1, row_[0], F12B, fill=GRAYF)
        put(ws4, r4, 2, row_[1], align=LEFT)
        put(ws4, r4, 3, row_[2], align=LEFT)
        ws4.row_dimensions[r4].height = 52
        r4 += 1
    if (delta or 0) < 0:
        hl_cell = (4, 2 if chg_1000 > 0 else 3)      # 净空扩大行
    elif (delta or 0) > 0:
        hl_cell = (3, 2 if chg_1000 > 0 else 3)      # 净空收窄行
    else:
        hl_cell = (5, 2)
    ws4.cell(row=hl_cell[0], column=hl_cell[1] + OFF).fill = YELLOWF
    r4 += 1
    put(ws4, r4, 1, "今日定位", F12B, fill=YELLOWF)
    put(ws4, r4, 2, gen_matrix, align=LEFT)
    merge(ws4, r4, 2, 3)
    ws4.row_dimensions[r4].height = fit_line(gen_matrix, CHARS_PER_LINE_2COL)
    r4 += 2
    put(ws4, r4, 1, "判读要点", F12B, fill=GRAYF)
    merge(ws4, r4, 1, 3)
    r4 += 1
    for t in ["同一组合放到不同位置含义相反：净空扩大+指数上涨若在高位，是主动看空/锁收益套保；若在低位，更多是反弹初段对冲盘的试探性设防，需看后续是否延续。",
              "幅度是过滤噪声的关键：单日变动占总净空的比例过小（如<5%）时属噪声级扰动，应以后续2-3个交易日净空方向为准。",
              "必须交叉验证：近月（%s/%s）代表短期方向性，季月（2703/2706）多为移仓与中期套保；只看合计曾误判（2026-09-17/18上涨中加空被当作被动对冲）。"
              % (CONTRACT, _next_month(CONTRACT)),
              "【被动套保 vs 主动做空】多头套保＝买现货＋开期指空单，因此指数下跌时不能轻易把净空扩大算作套保："
              "①维持对冲敞口所需手数在现货与期货同跌时基本不变（敞口与期货名义同步缩水），被动盘无需加空，"
              "实际加空超出该基准即为主动；②套保建仓必然伴随现货买入，须看主力资金——净流出＝现货在减仓而非建仓，"
              "与期货加空构成双重看空；③价比净额更可靠的是持仓量方向：价跌增仓＝新空进场主导，价跌减仓＝多头止损离场。",
              "上述三条已固化为固定流程，每个交易日必做：逐条结果见「性质鉴别」工作表（判据一 被动套保基准手数／"
              "判据二 持仓量方向／判据三 现货资金流），并在「当日摘要·五」中逐行展开；缺数据的判据照常列出并标注「无法判定」。"]:
        block(ws4, r4, t, 3, fit_line(t, 50))
        r4 += 1

    # ============ Sheet5 性质鉴别：被动套保 vs 主动做空（固定三条判据，逐日核查） ============
    ws5 = wb.create_sheet("性质鉴别")
    ws5.column_dimensions["A"].width = 3
    ws5.column_dimensions["B"].width = 3
    ws5.column_dimensions["C"].width = 18
    for col in "DEFG":
        ws5.column_dimensions[col].width = 28
    W5_PER_LINE = 62
    put(ws5, 1, 1, "净空变化的性质鉴别：被动套保 vs 主动做空（每日固定逐条核查）", F12B, fill=GRAYF, align=LEFT)
    merge(ws5, 1, 1, 5)
    r5 = 3
    put(ws5, r5, 1, "判别前提", F12B, fill=YELLOWF)
    merge(ws5, r5, 1, 5)
    r5 += 1
    block(ws5, r5, "多头套保＝买现货＋开期指空单，因此「指数下跌」与「买入现货」方向冲突——"
                   "下跌日的净空扩大一律先按主动加空处理，除非下列三条判据同时指向套保。", 5,
          _row_h("多头套保＝买现货＋开期指空单，因此「指数下跌」与「买入现货」方向冲突——下跌日的净空扩大一律先按主动加空处理，除非下列三条判据同时指向套保。", W5_PER_LINE))
    r5 += 2
    for i, (no, nm, basis) in enumerate(HEDGE_METHOD):
        _no, _nm, val, verdict = gen_diag["rows"][i]
        put(ws5, r5, 1, "%s · %s" % (no, nm), F12B, fill=GRAYF, align=LEFT)
        merge(ws5, r5, 1, 5)
        r5 += 1
        for tag, txt in (("口径", basis), ("今日实测", val), ("判定", verdict)):
            block(ws5, r5, "【%s】%s" % (tag, txt), 5, _row_h(txt, W5_PER_LINE))
            r5 += 1
        r5 += 1
    put(ws5, r5, 1, "综合定性", F12B, fill=YELLOWF)
    merge(ws5, r5, 1, 5)
    r5 += 1
    block(ws5, r5, gen_diag["verdict"], 5, _row_h(gen_diag["verdict"], W5_PER_LINE))
    r5 += 2
    put(ws5, r5, 1, "口径限制与使用注意", F12B, fill=GRAYF)
    merge(ws5, r5, 1, 5)
    r5 += 1
    block(ws5, r5, HEDGE_CAVEAT, 5, _row_h(HEDGE_CAVEAT, W5_PER_LINE))

    wb.save(out_path)


def _next_month(m):
    """合约月份 YYMM -> 下一个 YYMM"""
    y, mm = int(m[:2]), int(m[2:])
    mm += 1
    if mm > 12:
        y, mm = y + 1, 1
    return "%02d%02d" % (y, mm)


def make_analysis_file(date_obj, quotes, mkt, products, detail, old_products, settles,
                       series, aux, iv, total, delta, out_path, old_detail=None):
    """先生成 xlsx，再打 AIGC 标。"""
    build_analysis(date_obj, quotes, mkt, products, detail, old_products, settles,
                   series, aux, iv, total, delta, out_path, old_detail=old_detail)
    aigc_label(out_path)


# ==================== 6. 主流程 ====================
SKIP_FILE = os.path.join(DATA, "skipped_dates.json")


def catchup_missing(today_str):
    """回补追踪表中历史空缺行（C 列为空且日期早于今天）。

    可回补：B星期 / C~F四品种净持仓 / G合计公式 / H净变动；
            K基差、L~M涨跌幅（仅当 12 根K线窗口覆盖该日）。
    无法回补：I~J隐波（期权链接口只有当前快照，没有历史数据），留空。
    幂等：已填过的行跳过；确认休市的日期记入 skipped_dates.json 不再重试。
    """
    import openpyxl
    import warnings
    warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(TRACK, data_only=False)
    ws = wb.active
    if not _is_weekday_header(ws.cell(3, 2).value):
        return
    skipped = {}
    if os.path.exists(SKIP_FILE):
        try:
            skipped = json.load(open(SKIP_FILE, encoding="utf-8"))
        except Exception:
            skipped = {}
    missing = []
    for r in range(4, ws.max_row + 1):
        a = ws.cell(r, 1).value
        m = re.match(r"(\d{4})年(\d{1,2})月(\d{1,2})日", str(a or ""))
        if not m:
            continue
        ymd = "%s-%02d-%02d" % (int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if ymd >= today_str or ws.cell(r, 3).value not in (None, "") or ymd in skipped:
            continue
        missing.append((ymd, r))
    if not missing:
        log("历史回补检查：无缺失行")
        return
    log("发现 %d 个历史缺失行（%s ~ %s），开始回补…" %
        (len(missing), missing[0][0], missing[-1][0]))
    mkt = fetch_market(today_str)
    idx_close = {nm: {x["date"]: x["close"] for x in h} for nm, h in mkt["hist"].items()}
    im_close = {x["date"]: x["close"] for x in mkt["im"]["hist"]}

    done, n = 0, len(missing)
    for ymd, r in missing:
        done += 1
        log("回补进度 %d/%d：%s（表内第 %d 行）…" % (done, n, ymd, r))
        raw = fetch_positions(ymd)
        if not raw:
            log("  %s 无持仓数据（休市或数据未发布），记入跳过清单" % ymd)
            skipped[ymd] = "no_position"
            json.dump(skipped, open(SKIP_FILE, "w", encoding="utf-8"),
                      ensure_ascii=False, indent=1)
            continue
        products, detail, settles = calc_nets(raw)
        total = sum(products.values())
        log("  %s 净持仓: %s 合计=%d" % (ymd,
            " ".join("%s=%d" % (k, products[k]) for k in PRODUCTS), total))
        update_series(ymd, products)
        iv = {"call_iv": None, "put_iv": None, "basis": None,
              "chg_sh": None, "chg_1000": None}
        if ymd in idx_close.get("中证1000", {}) and ymd in im_close:
            iv["basis"] = round(idx_close["中证1000"][ymd] - im_close[ymd])
        for nm, key in (("上证指数", "chg_sh"), ("中证1000", "chg_1000")):
            h = mkt["hist"].get(nm, [])
            i = next((k for k, x in enumerate(h) if x["date"] == ymd), None)
            if i and i > 0:
                iv[key] = round(h[i]["close"] / h[i - 1]["close"] - 1, 6)
        if iv["basis"] is None:
            log("  %s 基差/涨跌幅超出K线窗口，留空；隐波无历史数据，留空" % ymd)
        else:
            log("  %s 隐波无历史数据留空；基差 %d 点已回补" % (ymd, iv["basis"]))
        day_obj = datetime.datetime.strptime(ymd, "%Y-%m-%d").date()
        status, row = update_tracking(day_obj, products, total, iv)
        log("  %s 写入结果: %s (行 %s)" % (ymd, status, row))
    log("历史回补结束")


def read_aux_from_track():
    import openpyxl
    import warnings
    warnings.filterwarnings("ignore")
    wb = openpyxl.load_workbook(TRACK, data_only=True)
    ws = wb.active
    aux = []
    for r in range(4, ws.max_row + 1):
        if ws.cell(r, 12).value is None:  # L 列沪指为空 = 未填
            continue
        aux.append(tuple(ws.cell(r, c).value for c in range(9, 14)))  # I J K L M，A列日期另外取
        aux[-1] = (str(ws.cell(r, 1).value or ""),) + aux[-1]
    return aux[-10:]


def main():
    ensure_log_bom()
    os.makedirs(OUTDIR, exist_ok=True)
    os.makedirs(ANALYSIS_DIR, exist_ok=True)
    now = datetime.datetime.now()
    date_obj = now.date()
    date_str = date_obj.strftime("%Y-%m-%d")
    log("==== 开始 %s ====" % date_str)
    ensure_weekday_fill()  # B 列星期空白补齐（幂等）

    # 1. 持仓
    raw = fetch_positions(date_str)
    if not raw:
        log("当日无持仓数据（休市或数据未发布），转入历史回补模式。")
        catchup_missing(date_str)
        log("==== 完成 ====")
        return 0
    products, detail, settles = calc_nets(raw)
    total = sum(products.values())
    log("净持仓: " + " ".join("%s=%d" % (k, products[k]) for k in PRODUCTS) + " 合计=%d" % total)

    # 2. 序列
    prev_date, prev_products = update_series(date_str, products)
    old_products = prev_products or {k: 0 for k in PRODUCTS}
    delta = total - sum(old_products.values())

    # 3. 行情
    mkt = fetch_market(date_str)
    # 抓取失败（网络/接口异常）时不得用空值冲掉本次已抓到的有效数据：
    # 同一天可能被多次运行，后一次失败不应覆盖前一次成功的字段。
    _mp = os.path.join(OUTDIR, "market_%s.json" % date_str)
    if os.path.exists(_mp):
        try:
            _old = json.load(open(_mp, encoding="utf-8"))
            for _k, _v in (("zdfb", mkt.get("zdfb")), ("ff_all", mkt.get("ff_all")),
                           ("breadth", mkt.get("breadth"))):
                if not _v and _old.get(_k):
                    mkt[_k] = _old[_k]
                    log("  %s 本次未取到，沿用已有数据" % _k)
        except Exception:
            pass
    json.dump(mkt, open(_mp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    quotes = build_quotes(mkt, date_str)
    im_close = mkt["im"]["close"]
    if im_close is None or mkt["im"]["date"] != date_str:
        log("警告：IM%s 当日收盘未取到（%s），基差与隐波将缺省" % (CONTRACT, mkt["im"]["date"]))

    # 4. 隐波 + 基差
    iv = {"call_iv": None, "put_iv": None, "basis": None, "chg_sh": quotes["上证指数"]["chg"],
          "chg_1000": quotes["中证1000"]["chg"]}
    if im_close:
        basis = round(quotes["中证1000"]["close"] - im_close)
        iv["basis"] = basis
        try:
            iv = fetch_iv(quotes["中证1000"]["close"], im_close)
            iv["basis"] = basis
            iv["chg_sh"] = quotes["上证指数"]["chg"]
            iv["chg_1000"] = quotes["中证1000"]["chg"]
            json.dump(iv, open(os.path.join(OUTDIR, "iv_%s.json" % date_str), "w", encoding="utf-8"),
                      ensure_ascii=False, indent=1)
            log("隐波: 认购 %.2f%% / 认沽 %.2f%% | 基差 %d 点" % (iv["call_iv"], iv["put_iv"], basis))
        except Exception as e:
            log("隐波反解失败（%s），本日隐波将留空" % type(e).__name__)
    json.dump({"date": date_str, "products": products, "total": total, "detail": detail},
              open(os.path.join(OUTDIR, "position_%s.json" % date_str), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # 5. 更新追踪表
    status, row = update_tracking(date_obj, products, total, iv)
    log("追踪表: %s (行 %s)" % (status, row))

    # 6. 分析文件
    aux = read_aux_from_track()
    old_detail = None
    if prev_date:
        p = os.path.join(OUTDIR, "position_%s.json" % prev_date)
        if os.path.exists(p):
            old_detail = json.load(open(p, encoding="utf-8")).get("detail")
    out = os.path.join(ANALYSIS_DIR, "%s盘面多维分析.xlsx" % date_str)
    make_analysis_file(date_obj, quotes, mkt, products, detail, old_products, settles,
                       series=json.load(open(SERIES, encoding="utf-8")), aux=aux, iv=iv,
                       total=total, delta=delta, out_path=out, old_detail=old_detail)
    log("分析文件: " + out)

    # 7. 历史回补（幂等：只补 C 列为空的历史行）
    catchup_missing(date_str)
    log("==== 完成 ====")
    return 0


if __name__ == "__main__":
    sys.exit(main())
