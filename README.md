# 股指期货持仓追踪 · 安卓版

把桌面上的 `daily_job.py` 搬到手机上：每交易日自动抓取持仓、行情、期权隐波，
生成与电脑端**完全一致**的盘面多维分析（含结论、性质鉴别三条判据、次日观察位）。

## 它和电脑端是什么关系

手机里跑的就是主项目那份 `daily_job.py`（同一份文件复制来的），
抓取、计算、判据、结论全部复用，只是把数据目录重定向到手机私有存储。
**不存在两套逻辑**，所以两端的数字和口径永远一致。

与电脑端的两点必要差异：

| 项目 | 电脑端 | 手机端 |
|---|---|---|
| 结果载体 | `盘面多维分析/日期.xlsx` | APP 界面直接展示 + 同目录 HTML（便于分享） |
| 追踪主表 | 维护 `国内股指期货持仓追踪.xlsx` | 不维护（手机上打不开也没必要） |
| 昨日基差/隐波 | 从主表读 | 从当日存档构造（等价） |

## 目录结构

```
main.py            Kivy 界面（展示 + 刷新）
app_job.py         流程编排：抓数 → 计算 → 生成结论 → 落盘
daily_job.py       核心逻辑（与主项目同源，用 sync_core.py 同步）
service.py         后台定时服务（APP 退出后仍可定时跑）
seed_series.json   预置净持仓序列底稿（决定净变动与历史趋势）
seed_outputs/      预置最近几日持仓明细+隐波存档（判据二、昨日对比要用）
buildozer.spec     APK 打包配置
sync_core.py       从主项目同步最新 daily_job.py 与预置数据
.github/           GitHub Actions 自动编译
```

## 编译 APK（本机编不了，用 GitHub Actions）

这台电脑没有 JDK / Android SDK，装不了打包工具链，所以走云端编译：

> 本目录已经帮你 `git init` 并提交过了（23 个文件，`main` 分支），
> 所以你只需要建仓库 + 推送，不用再手动 add/commit。

1. 在 GitHub 新建**空**仓库（Public/Private 都行，不要勾选 README/.gitignore/license）。

   **方式 A（推荐）**：双击 `推送到GitHub.bat`，粘贴仓库地址回车，
   按提示登录即可（会弹出 GitHub 登录窗口）。

   **方式 B（命令行）**：

   ```bash
   git remote add origin https://github.com/<你的账号>/<仓库名>.git
   git push -u origin main
   ```

2. 推完后**通常会自动开始编译**（工作流已配置 push 到 main 触发）。
   若没触发，去仓库页面 → **Actions** → 左侧选 `Build Android APK`
   → 右侧 **Run workflow** → 点绿色按钮。

3. 等待约 20～40 分钟（首次会下载 Android SDK/NDK，约 1.5GB）。跑完后在
   该次运行页面底部 **Artifacts** 里下载 `index-futures-tracker-apk`，解压得到 `.apk`。

4. 把 APK 传到手机安装。首次安装会提示"未知来源"，允许即可。

> 之后每次改代码，push 后手动 Run workflow 就会出新版。
> 也可以打 tag 触发：`git tag v1.1 && git push --tags`

## 使用

- 打开 APP 即显示最近一次分析；若当天还没跑过会自动开始抓取（约 1 分钟，需联网）。
- 顶部两个按钮：**立即刷新** / **自动刷新开关**。
- 刷新策略做了三层，尽量保证"每个交易日都有内容"：
  1. 启动即检查：本地没有当日完整数据就自动跑；
  2. 前台轮询：APP 开着时每 10 分钟检查一次；
  3. 后台服务：APP 退出后每 15 分钟检查一次（需挂前台通知才活得久）。

**装完后建议做一次设置**（否则后台可能被系统杀掉，导致当天没数据）：
- 系统设置 → 应用 → 本 APP → 电池 → 选"无限制/不优化"；
- 允许通知（后台服务靠常驻通知存活）；
- 若手机有"自启动/关联启动"管理，把本 APP 加进白名单。

## 数据存在哪

`/data/data/org.example.indexfuturestracker/files/app/`
（APP 私有目录，卸载即清空；不需要任何存储权限，也不会污染相册/下载目录）

- `data/em_position_series.json` 净持仓序列（首次由 seed 初始化，之后每天累加）
- `outputs/latest.json` 最近一次分析结果，界面直接读它
- `outputs/analysis_<日期>.json` 按日归档
- `盘面多维分析/<日期>盘面多维分析.html` 可分享的 HTML

## 电脑端改了逻辑之后

```bash
python sync_core.py     # 同步 daily_job.py + 最新序列/存档
git add . && git commit -m "sync" && git push
```
然后重新 Run workflow 编译、安装新版。

## 已知限制

- **期权/资金接口只在工作日开盘时段有数据**。周末打开 APP 会自动往前找最近交易日，
  但重抓的期权价格可能与交易日收盘时不同（周末返回的是最新价而非当日结算价），
  所以周末看到的隐波可能和交易日存档有出入——以交易日当天跑出来的为准。
- 主力资金接口（东财 push2）在部分网络下取不到，此时该字段显示"—"，
  判据三会如实标注"无法交叉验证"，不影响其余结论。
- APK 体积约 40～60MB（含 Python 解释器与 Kivy）。
- 后台定时受安卓省电策略影响，国产机型尤其严格；若发现当天没自动更新，
  检查上面的电池优化设置，或打开 APP 点一次"立即刷新"。
