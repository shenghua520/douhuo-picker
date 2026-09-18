---
feature: brand-category-scrape
status: designed
updated: 2026-09-10
branch: feat/scrape-choice
commits:
---

# 品牌品类总表批量抓取

## Report

## [S1] Problem
根据 `总表_品牌品类汇总.xlsx` 的品牌/大类/子类，批量检索抖货商品并导出可人工筛选的 Excel。

## [S2] Design
- 输入：总表 sheet「总表」列 品牌名称/大类/子类（约 101 行，约 40+ 品牌）。
- 检索：`getGoodsList` `type=1, spu_name=品牌名` 翻页；标题需包含品牌；**排除京东**（supply_type=1 或 supply_name 含京东）。
- 归类：写入品牌在总表中的大类/子类；若标题命中子类关键词则记「匹配子类」。
- 去重合并：同 `spu_name`（规范化）+ `default_sku` 为一组；组内与结果均按 `plat_price` 降序；保留 `重复组键/组内goods_ids`。
- 表头：第1行英文键、第2行中文；另建「字段说明」表（含 is_freight 1=包邮 等取值）。
- 分表：按大类多个 sheet（油/米面/个护/乳品/饮料/纸品清洁…）。
- 稳定性：登录后额外等待页面加载；每品牌/每5品牌输出进度与 ETA，并写 progress_status.txt。
- CLI：`python douhuo_scraper/brand_category.py [--brands-limit N] [--max-per-brand N]`

## [S3] Out of Scope
- 不对每条商品调 getGoodsDetail（避免极慢）。
- 不改总表源文件。

## Tasks
- [ ] T1: 读总表并品牌检索（去京东）— acceptance: 抽样品牌能返回非京东商品 (covers: S2)
- [ ] T2: 分组排序分表导出 — acceptance: 多大类 sheet + 字段说明 + 组内 goods_ids (covers: S2)
- [ ] T3: 全量跑批 — acceptance: 输出 xlsx 且行数>0 (covers: S2)
