# -*- coding: utf-8 -*-
"""把主项目里最新的 daily_job.py 同步到 APP 工程。

电脑端 daily_job.py 改了（口径调整、bug 修复等）之后跑一次这个脚本，
再重新提交+编译，手机端才会跟着变——两端跑的是同一份分析逻辑。

用法：python sync_core.py
"""

import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.normpath(os.path.join(HERE, "..", "daily_job.py"))
DST = os.path.join(HERE, "daily_job.py")


def main():
    if not os.path.exists(SRC):
        print("未找到主项目 daily_job.py：%s" % SRC)
        print("（若本工程已脱离主项目独立存在，请手动复制 daily_job.py 到本目录）")
        return 1
    shutil.copy(SRC, DST)
    print("已同步：%s -> %s" % (SRC, DST))
    # 顺带同步历史序列底稿与最近存档，保证新装 APP 开箱即有趋势和判据二所需数据
    seed = os.path.normpath(os.path.join(HERE, "..", "data", "em_position_series.json"))
    if os.path.exists(seed):
        shutil.copy(seed, os.path.join(HERE, "seed_series.json"))
        print("已同步序列底稿 seed_series.json")
    outs = os.path.normpath(os.path.join(HERE, "..", "outputs"))
    seed_out = os.path.join(HERE, "seed_outputs")
    if os.path.isdir(outs):
        os.makedirs(seed_out, exist_ok=True)
        n = 0
        for fn in sorted(os.listdir(outs))[-14:]:   # 最近若干交易日
            if fn.startswith(("position_", "iv_")) and fn.endswith(".json"):
                shutil.copy(os.path.join(outs, fn), os.path.join(seed_out, fn))
                n += 1
        print("已同步最近存档 %d 个文件到 seed_outputs/" % n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
