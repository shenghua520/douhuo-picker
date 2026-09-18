"""油类全量商品：名称解析 品牌/子类/规格 → 分表 Excel。

进度：每品牌/关键词打点；每约 2 分钟输出进度+ETA；同步 progress_status.txt。
"""

from __future__ import annotations

import json
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from client import DouhuoClient, NotLoggedIn  # noqa: E402
from excelio import FIELD_ZH, flatten_value, union_keys  # noqa: E402

PROFILE = Path(ROOT / "state" / "browser_profile")
OUT_DIR = ROOT / "output" / "油类"
STATUS_FILE = ROOT / "output" / "progress_status.txt"
OUT_XLSX = OUT_DIR / "油类全量商品_品牌子类规格.xlsx"

HEADER_FILL = PatternFill("solid", fgColor="1D1F27")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
LABEL_FILL = PatternFill("solid", fgColor="3D4450")
LABEL_FONT = Font(color="FFFFFF", size=10)

# 采集关键词（油品）
OIL_KEYWORDS = [
    "菜籽油", "玉米油", "花生油", "大豆油", "调和油", "橄榄油", "亚麻籽油",
    "葵花籽油", "山茶油", "核桃油", "稻米油", "米糠油", "芝麻油", "香油",
    "茶籽油", "芥花油", "红花籽油", "葡萄籽油", "小麦胚芽油", "椰子油",
    "棕榈油", "猪油", "牛油", "鸭油", "食用油", "植物油", "色拉油",
    "食用调和油", "低芥酸菜籽油", "非转基因", "压榨油",
]

# 品牌词典（用于从名称提取品牌，长词优先）
BRANDS = [
    "金龙鱼", "福临门", "鲁花", "道道全", "胡姬花", "多力", "长寿花",
    "西王", "刀唛", "日清", "鹰唛", "骆驼唛", "刀唛", "红灯", "元宝",
    "鲤鱼", "香满园", "厨宝", "海皇", "狮球唛", "刀唛", "花旗",
    "欧丽薇兰", "贝蒂斯", "品利", "安达露西", "阿格利司", "皇家蒙特垒",
    "融氏", "山萃", "润心", "金浩", "贵太太", "老树根", "绿宝",
    "长寿花", "龙大", "喜燕", "五湖", "中粮", "中储粮", "九三",
    "北大荒", "十月稻田", "福花", "纳福汇", "日恋", "日日恋",
    "鲁花", "齐云山", "野岭", "接福", "老史家", "金沙河",
    "陈克明", "裕湘", "金健", "克明", "想念", "博大",
]

# 子类关键词（长优先）
SUBCATS = [
    "低芥酸菜籽油", "非转基因菜籽油", "物理压榨菜籽油", "浓香菜籽油",
    "菜籽油", "菜油", "芥花油", "玉米油", "玉米胚芽油",
    "花生油", "古法花生油", "浓香花生油",
    "大豆油", "色拉油", "食用调和油", "调和油",
    "橄榄油", "特级初榨橄榄油", "初榨橄榄油",
    "亚麻籽油", "亚麻油", "核桃油", "山茶油", "茶籽油",
    "葵花籽油", "葵花油", "稻米油", "米糠油", "芝麻油", "小磨香油", "香油",
    "红花籽油", "葡萄籽油", "小麦胚芽油", "椰子油", "棕榈油",
    "猪油", "牛油", "鸭油", "食用油", "植物油", "粮油",
]

VALUE_LEGEND = [
    ("品牌 brand", "从商品名称解析的品牌；未识别为空"),
    ("子类 oil_type", "从名称解析的油品子类，如菜籽油/玉米油"),
    ("规格 spec", "从名称解析的规格，如 5L / 1.8L / 900ml"),
    ("is_freight / 是否包邮", "1=包邮；2=不包邮；空=未知"),
    ("status / 上下架状态", "1=在售；0=售罄/下架"),
    ("supply_type / 渠道类型", "1=京东；5=厂商特卖；7=微唯宝特卖；17=华东一仓；19=云采渠道"),
    ("plat_price / 代发价", "平台代发成本价（元）"),
    ("retail_price / 零售价", "建议零售价（元）"),
    ("profit / 利润率(%)", "平台利润率百分比"),
]

PREFERRED = [
    "品牌", "子类", "规格",
    "goods_id", "spu_sn", "spu_name", "default_sku",
    "supply_type", "supply_name", "plat_price", "retail_price", "profit",
    "min_reduce", "max_reduce", "is_freight", "status", "sort",
    "main_img", "is_update", "update_time",
    "匹配关键词", "商品链接", "采集时间",
]

ZH_EXTRA = {
    "品牌": "品牌 brand",
    "子类": "子类 oil_type",
    "规格": "规格 spec",
    "匹配关键词": "命中关键词",
}


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def write_status(text: str) -> None:
    try:
        STATUS_FILE.parent.mkdir(parents=True, exist_ok=True)
        STATUS_FILE.write_text(text + "\n", encoding="utf-8")
    except OSError:
        pass


def parse_brand(name: str) -> str:
    n = name or ""
    for b in sorted(BRANDS, key=len, reverse=True):
        if b in n:
            return b
    return ""


def parse_subcat(name: str) -> str:
    n = name or ""
    for s in sorted(SUBCATS, key=len, reverse=True):
        if s in n:
            return s
    return ""


def parse_spec(name: str) -> str:
    """提取体积/重量规格，归一化为 如 5L / 500ml / 2.5kg。"""
    n = name or ""
    # 优先匹配 数字+单位
    pats = [
        r"(\d+(?:\.\d+)?)\s*[Ll升](?![a-zA-Z])",
        r"(\d+(?:\.\d+)?)\s*[Mm][Ll]",
        r"(\d+(?:\.\d+)?)\s*[Kk][Gg]",
        r"(\d+(?:\.\d+)?)\s*[Gg](?![a-zA-Z])",
        r"(\d+(?:\.\d+)?)\s*斤",
    ]
    for p in pats:
        m = re.search(p, n)
        if not m:
            continue
        num = m.group(1)
        unit_raw = m.group(0)
        if re.search(r"[Mm][Ll]", unit_raw):
            return f"{num}ml"
        if re.search(r"[Kk][Gg]", unit_raw):
            return f"{num}kg"
        if re.search(r"斤", unit_raw):
            return f"{num}斤"
        if re.search(r"[Gg]", unit_raw):
            return f"{num}g"
        return f"{num}L"
    return ""


def is_oil_name(name: str) -> bool:
    n = name or ""
    keys = ["油", "食用油", "植物油"]
    return any(k in n for k in keys)


def fetch_keyword(client: DouhuoClient, kw: str) -> list[dict]:
    out = []
    page = 1
    while True:
        res = client.goods_list(page=page, limit=100, **{"type": 1, "spu_name": kw})
        items = res.get("list") or []
        pages = res.get("pages")
        if not items:
            break
        out.extend(items)
        if pages is not None and page >= int(pages):
            break
        page += 1
        if page % 5 == 0:
            log(f"    「{kw}」 page={page} 累计 {len(out)}…")
    return out


def write_excel(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    headers = union_keys(rows, PREFERRED)
    wb = Workbook()
    ws0 = wb.active
    ws0.title = "字段说明"
    ws0.append(["字段 Field", "说明 Description"])
    for c in ws0[1]:
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
    for k, v in VALUE_LEGEND:
        ws0.append([k, v])
    for h in headers:
        ws0.append([h, ZH_EXTRA.get(h) or FIELD_ZH.get(h) or h])
    ws0.column_dimensions["A"].width = 28
    ws0.column_dimensions["B"].width = 48

    ws = wb.create_sheet("油类全量")
    for col, key in enumerate(headers, start=1):
        c1 = ws.cell(1, col, key)
        c1.fill = HEADER_FILL
        c1.font = HEADER_FONT
        c2 = ws.cell(2, col, ZH_EXTRA.get(key) or FIELD_ZH.get(key) or key)
        c2.fill = LABEL_FILL
        c2.font = LABEL_FONT
        ws.column_dimensions[get_column_letter(col)].width = min(max(len(key) + 2, 12), 28)
    for r_i, row in enumerate(rows, start=3):
        for c_i, key in enumerate(headers, start=1):
            ws.cell(r_i, c_i, flatten_value(row.get(key)))
    ws.freeze_panes = "A3"
    if rows:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{ws.max_row}"

    # 按子类分表
    by_sub: dict[str, list[dict]] = {}
    for r in rows:
        by_sub.setdefault(r.get("子类") or "未识别子类", []).append(r)
    for sub, items in sorted(by_sub.items(), key=lambda x: -len(x[1])):
        safe = re.sub(r'[\\/*?:\[\]]', "_", sub)[:31]
        ws2 = wb.create_sheet(safe[:31])
        for col, key in enumerate(headers, start=1):
            c1 = ws2.cell(1, col, key)
            c1.fill = HEADER_FILL
            c1.font = HEADER_FONT
            c2 = ws2.cell(2, col, ZH_EXTRA.get(key) or FIELD_ZH.get(key) or key)
            c2.fill = LABEL_FILL
            c2.font = LABEL_FONT
            ws2.column_dimensions[get_column_letter(col)].width = min(max(len(key) + 2, 12), 28)
        for r_i, row in enumerate(items, start=3):
            for c_i, key in enumerate(headers, start=1):
                ws2.cell(r_i, c_i, flatten_value(row.get(key)))
        ws2.freeze_panes = "A3"
    wb.save(path)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    last_progress = t0
    seen: set[str] = set()
    rows: list[dict] = []

    with DouhuoClient(PROFILE, headless=True) as client:
        client.open_entry()
        client._page.wait_for_timeout(4000)
        log(f"登录态: {client.login_state()}")
        log(f"油品关键词 {len(OIL_KEYWORDS)} 个，开始采集")

        for ki, kw in enumerate(OIL_KEYWORDS, 1):
            try:
                goods = fetch_keyword(client, kw)
            except NotLoggedIn:
                log("登录失效，请重新 --login")
                return 3
            except Exception as exc:  # noqa: BLE001
                log(f"关键词「{kw}」失败: {exc}")
                goods = []
            new = 0
            for item in goods:
                name = str(item.get("spu_name") or "")
                gid = str(item.get("goods_id") or "")
                if not gid or gid in seen:
                    continue
                if not is_oil_name(name):
                    continue
                seen.add(gid)
                brand = parse_brand(name)
                sub = parse_subcat(name)
                spec = parse_spec(name)
                row = dict(item)
                row["品牌"] = brand
                row["子类"] = sub
                row["规格"] = spec
                row["匹配关键词"] = kw
                row["商品链接"] = f"https://www.douhuomall.com/choice/v3/#/goods?id={gid}&nav_type=1"
                row["采集时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                rows.append(row)
                new += 1
            log(f"[{ki}/{len(OIL_KEYWORDS)}] 「{kw}」接口 {len(goods)} → 油类新增 {new} 累计 {len(rows)}")

            now = time.time()
            if now - last_progress >= 300 or ki == len(OIL_KEYWORDS):
                elapsed = now - t0
                rate = ki / elapsed if elapsed else 0
                remain_kw = len(OIL_KEYWORDS) - ki
                eta_min = remain_kw / rate / 60 if rate else 0
                msg = (
                    f"[进度] 关键词 {ki}/{len(OIL_KEYWORDS)} "
                    f"({ki*100//len(OIL_KEYWORDS)}%) 累计油类商品 {len(rows)} "
                    f"已用 {elapsed/60:.1f}分钟 ETA≈{eta_min:.0f}分钟"
                )
                log(msg)
                write_status(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")
                last_progress = now

    # 排序：子类 → 品牌 → 代发价降序
    def sort_key(r):
        try:
            price = -float(r.get("plat_price") or 0)
        except (TypeError, ValueError):
            price = 0
        return (str(r.get("子类") or "zzz"), str(r.get("品牌") or "zzz"), price)

    rows.sort(key=sort_key)
    write_excel(OUT_XLSX, rows)

    # Review 统计
    total = len(rows)
    has_brand = sum(1 for r in rows if r.get("品牌"))
    has_sub = sum(1 for r in rows if r.get("子类"))
    has_spec = sum(1 for r in rows if r.get("规格"))
    log("=" * 50)
    log(f"完成 total={total} 品牌识别 {has_brand}({has_brand*100//max(total,1)}%) "
        f"子类 {has_sub}({has_sub*100//max(total,1)}%) 规格 {has_spec}({has_spec*100//max(total,1)}%)")
    sub_cnt: dict[str, int] = {}
    brand_cnt: dict[str, int] = {}
    for r in rows:
        sub_cnt[r.get("子类") or "空"] = sub_cnt.get(r.get("子类") or "空", 0) + 1
        brand_cnt[r.get("品牌") or "空"] = brand_cnt.get(r.get("品牌") or "空", 0) + 1
    log("子类分布 Top15: " + ", ".join(f"{k}:{v}" for k, v in sorted(sub_cnt.items(), key=lambda x: -x[1])[:15]))
    log("品牌分布 Top15: " + ", ".join(f"{k}:{v}" for k, v in sorted(brand_cnt.items(), key=lambda x: -x[1])[:15]))
    log(f"输出 {OUT_XLSX}")
    write_status(f"[{datetime.now().strftime('%H:%M:%S')}] 完成 {total} 条 → {OUT_XLSX.name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
