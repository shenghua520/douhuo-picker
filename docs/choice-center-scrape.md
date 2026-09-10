# 选品中心五表全量抓取 / Choice Center Full Scrape

> 中英双语 · Bilingual (ZH / EN)

## 中文

### 是什么

在原有「Excel 驱动关键词选品」之外，新增 `douhuo_scraper/`：一键抓取选品中心首页五大分区，导出 5 份 Excel。

| 输出文件 | 内容 | 实测量级 |
|---|---|---|
| `01_首页轮播图商品.xlsx` | banner 元数据 + 侧栏商品 + **轮播跳转专区全量商品** | ~744 行 |
| `02_首页每日必看专区商品.xlsx` | 每日必看各专区入口 + 商品 | ~5.7k 行 |
| `03_团购爆品轮播图商品.xlsx` | 首页团购爆品轮播位商品 | 数条 |
| `04_团购爆品所有商品.xlsx` | 团购爆品专栏全量 | ~7.9k 行 |
| `05_海量优品全部商品.xlsx` | 海量优品/全库全量 | ~12.5 万行 |

### 用法

```bash
cd douhuo_scraper

# 1) 首次：有头登录（扫码/短信），cookies 会保存
python run.py --login

# 2) 抓取（可分区）
python run.py --scrape banner,daily,tuan_top,tuan_all,youpin --fresh
# 或全部
python run.py --scrape all --fresh

# 3) 仅从 JSONL 重导 Excel
python run.py --export-only
```

### 表头约定

- 第 1 行：英文接口字段键（程序友好）
- 第 2 行：中文说明（人读友好）
- 数据从第 3 行开始，冻结至 A3

### 进度反馈

长任务（如 12 万条）每约 200 条或 20 秒打印进度，并写入：

```
output/progress_status.txt
```

### 权限说明

- 游客/试用账号无法访问「海量优品」及多数每日必看专栏（接口返回 1014）。
- 请先 `python run.py --login` 使用正式账号。

### 接口映射（逆向 + 实测）

| 分区 | API |
|---|---|
| 首页轮播图 | `getBannerImg` + `getTopSpecial` + 按 `link_url` 翻页 `getGoodsList` |
| 每日必看 | `getRecsSpecial` → 各 nav `getGoodsList` |
| 团购爆品轮播 | `getTopSpecial` 中 `nav_name=团购爆品` |
| 团购爆品全量 | `getGoodsList` `nav_id=11` |
| 海量优品 | `getGoodsList` `nav_id=7`（无权限时回退全站） |

---

## English

### What it is

Besides the original Excel-driven keyword picker, this repo adds `douhuo_scraper/`: scrape five homepage sections of the Choice Center into five Excel workbooks.

### Usage

```bash
cd douhuo_scraper
python run.py --login
python run.py --scrape all --fresh
python run.py --export-only
```

### Header layout

- Row 1: raw API field keys
- Row 2: Chinese labels
- Data starts at row 3 (freeze panes A3)

### Progress

Long jobs log every ~200 rows / 20s and also write `output/progress_status.txt`.

### Permissions

Tourist/trial accounts get HTTP business code **1014** on 海量优品 and most daily special navs. Use a formal account via `--login`.
