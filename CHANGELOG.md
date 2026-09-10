# 变更日志

项目所有值得注意的变更都记录在此。格式参考 [Keep a Changelog](https://keepachangelog.com/)。

## [Unreleased] · 2026-09-10

### Added
- 新增 `douhuo_scraper/`：选品中心五表全量抓取（首页轮播图/每日必看/团购爆品轮播/团购爆品全量/海量优品）。
- 轮播图 `link_url` 跳转专区商品一并采集；Excel 双行表头（英文字段 + 中文说明）。
- 长任务定时进度反馈（条数 / 百分比 / ETA）+ `output/progress_status.txt`。
- 双语文档：`docs/choice-center-scrape.md`。

## [Unreleased] · 2026-09-01

### Added
- 首版开源：抖货商城 Excel 驱动的自动选品工具。
- 核心能力：Playwright 持久化登录 + 页面内 fetch 调真实接口、openpyxl 读写任务表/结果表、批量处理 + 任务重跑 + 重复判重。
- 派生任务表：自动按品类合并去重，支持两种原始表头（`品类/品牌` 与 `品牌/系列`）。
- 名称 → ID 解析：类目/品牌/渠道列可直接填中文名或数字 ID。
- 本地重排：服务端排序不可靠时（利润率等计算字段），抓完按用户设定的排序字段本地重排。
- 就地校正工具 `dedup_existing.py`：同任务内双键（`商品ID` + `名称+双价`）判重并删除。
- 跑批前自动备份（`*_备份_<时间戳>.xlsx`），目标文件被 Excel 占用时给出明确错误而非静默丢数据。
- 试用账号自动识别：`getUserInfo.sys_id==11` 时打印警告，不静默写脏数据。
- GitHub Actions 冒烟测试：跑 `--init --dry-run` 验证任务表派生与脚本可执行。
- CHANGELOG（本文件）。

### Known limitations
- 仅在 Python 3.13.12 + Playwright Chromium 上验证；其它版本可跑但未做矩阵测试。
- 试用账号数据有限，重要场景请先用 `python run.py --login` 登录正式账号。
- 当前依赖抖货商城前端 JS bundle 路径稳定，前端大改版时需逆向重写。
