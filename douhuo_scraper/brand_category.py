"""按总表「品牌/大类/子类」批量抓取商品 → 按大类分表 Excel。

约定：
- 排除京东渠道（supply_type==1 / supply_name 含「京东」）
- 同名同 SKU 多 goods_id 归为一组，按代发价降序排列
- 双行表头 + 字段说明表
- 定时进度反馈 + progress_status.txt
- 打开页面后充分等待再采集，降低价格未就绪概率
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import time
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from client import API, DouhuoClient, NotLoggedIn  # noqa: E402
from excelio import FIELD_ZH, flatten_value, union_keys, zh_label  # noqa: E402

DEFAULT_PROFILE = ROOT / "state" / "browser_profile"
DEFAULT_INPUT = Path(r"C:\Users\shenghua\Desktop\数据\总表_品牌品类汇总.xlsx")
OUT_DIR = ROOT / "output" / "品牌品类"
JSONL_DIR = OUT_DIR / "jsonl"
STATUS_FILE = ROOT / "output" / "progress_status.txt"

HEADER_FILL = PatternFill("solid", fgColor="1D1F27")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
LABEL_FILL = PatternFill("solid", fgColor="3D4450")
LABEL_FONT = Font(color="FFFFFF", size=10)

JD_SUPPLY_TYPE = 1
EXCLUDE_SUPPLY_NAMES = ("京东",)

VALUE_LEGEND = [
    ("is_freight / 是否包邮", "1=包邮；2=不包邮；空/其它=未知"),
    ("status / 上下架状态", "1=在售；0=售罄/下架"),
    ("is_select / 是否已选", "1=已加入选品库；0=未选"),
    ("is_test / 是否测试", "1=测试商品；0=正式"),
    ("is_update / 是否更新中", "1=更新中（价格可能变动）；0=正常"),
    ("supply_type / 渠道类型", "1=京东；5=厂商特卖；7=微唯宝特卖；17=华东一仓；19=云采渠道"),
    ("plat_price / 代发价(发货价)", "平台代发成本价，单位元"),
    ("retail_price / 零售价", "建议零售价，单位元"),
    ("profit / 利润率(%)", "按平台公式计算的百分比利润率"),
    ("min_reduce/max_reduce / 集采价区间", "集采（批量采购）价上下限"),
    ("sort / 推荐指数", "平台推荐分，越大越靠前（负值按前端归零理解）"),
    ("default_sku / 默认SKU", "该 SPU 默认规格 ID"),
]


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def write_status(text: str) -> None:
    try:
        STATUS_FILE.parent.mkdir(parents=True, exist_ok=True)
        STATUS_FILE.write_text(text + "\n", encoding="utf-8")
    except OSError:
        pass


def norm_text(s: Any) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"\s+", "", s)
    return s.lower()


def is_jd(item: dict) -> bool:
    if int(item.get("supply_type") or 0) == JD_SUPPLY_TYPE:
        return True
    name = str(item.get("supply_name") or "")
    return any(x in name for x in EXCLUDE_SUPPLY_NAMES)


def load_tasks(path: Path) -> list[dict]:
    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["总表"]
    tasks: list[dict] = []
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1:
            continue
        brand = str(row[0] or "").strip()
        major = str(row[1] or "").strip()
        sub = str(row[2] or "").strip()
        if not brand or brand in {"品牌名称"}:
            continue
        tasks.append({"品牌名称": brand, "大类": major, "子类": sub})
    wb.close()
    return tasks


def group_key(item: dict) -> tuple[str, str]:
    name = norm_text(item.get("spu_name"))
    sku = str(item.get("default_sku") or "").strip()
    return name, sku


def match_subcategory(spu_name: str, brand: str, subs: list[str]) -> str:
    """在商品名中匹配子类关键词；匹配不到则返回空（仍保留，由大类表收录）。"""
    n = norm_text(spu_name)
    b = norm_text(brand)
    if b and b in n:
        pass
    # 子类关键词：去掉「大米-」前缀里的细分，用更长片段优先
    candidates = sorted({s for s in subs if s}, key=len, reverse=True)
    for sub in candidates:
        keys = [sub]
        if "-" in sub:
            keys.extend(sub.split("-"))
        for k in keys:
            k2 = norm_text(k)
            if k2 and k2 in n:
                return sub
    return ""


PREFERRED = [
    "大类", "子类", "品牌名称", "匹配子类",
    "goods_id", "spu_sn", "spu_name", "default_sku",
    "supply_type", "supply_name", "plat_price", "retail_price", "profit",
    "min_reduce", "max_reduce", "is_freight", "status", "sort",
    "main_img", "is_select", "is_test", "is_update", "update_time",
    "重复组键", "组内商品数", "组内goods_ids", "商品链接", "采集时间",
]


def write_sheet(ws: Worksheet, rows: list[dict], preferred: list[str] | None = None) -> None:
    headers = union_keys(rows, preferred)
    for col, name in enumerate(headers, start=1):
        c1 = ws.cell(row=1, column=col, value=name)
        c1.fill = HEADER_FILL
        c1.font = HEADER_FONT
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c2 = ws.cell(row=2, column=col, value=zh_label(name))
        c2.fill = LABEL_FILL
        c2.font = LABEL_FONT
        c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        width = max(len(name) + 2, len(str(zh_label(name))) * 2 + 2, 12)
        ws.column_dimensions[get_column_letter(col)].width = min(width, 36)
    for r_idx, row in enumerate(rows, start=3):
        for c_idx, key in enumerate(headers, start=1):
            ws.cell(row=r_idx, column=c_idx, value=flatten_value(row.get(key)))
    ws.freeze_panes = "A3"
    if rows:
        ws.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{ws.max_row}"


def write_workbook(path: Path, sheets: dict[str, list[dict]], legend: list[tuple[str, str]] | None = None) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    # 字段说明
    ws0 = wb.active
    ws0.title = "字段说明"
    ws0.append(["字段 Field", "说明 Description / 取值"])
    for cell in ws0[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
    for k, v in legend or VALUE_LEGEND:
        ws0.append([k, v])
    for k, v in FIELD_ZH.items():
        if k in PREFERRED:
            ws0.append([k, v])
    ws0.column_dimensions["A"].width = 36
    ws0.column_dimensions["B"].width = 60

    for name, rows in sheets.items():
        safe = re.sub(r'[\\/*?:\[\]]', "_", name)[:31] or "未命名"
        ws = wb.create_sheet(safe)
        write_sheet(ws, rows, PREFERRED)
    wb.save(path)


def search_brand_goods(client: DouhuoClient, brand: str, max_items: int | None = None) -> tuple[list[dict], Any]:
    """按商品名含品牌关键字检索并翻页。"""
    out: list[dict] = []
    page = 1
    total: Any = None
    payload_extra = {"type": 1, "spu_name": brand}
    while True:
        res = client.goods_list(page=page, limit=100, **payload_extra)
        items = res.get("list") or []
        pages = res.get("pages")
        total = res.get("total")
        if not items:
            break
        for item in items:
            name = str(item.get("spu_name") or "")
            # 品牌需出现在标题中，降低「福临门」命中无关礼盒噪声
            if norm_text(brand) not in norm_text(name):
                continue
            if is_jd(item):
                continue
            out.append(item)
            if max_items is not None and len(out) >= max_items:
                return out, total
        if pages is not None and page >= int(pages):
            break
        page += 1
        if page % 5 == 0:
            log(f"    品牌「{brand}」 page={page} 已筛 {len(out)} 条…")
    return out, total


def postprocess(rows: list[dict]) -> list[dict]:
    """同名同SKU分组 + 组内/整体按代发价降序。"""
    groups: dict[tuple[str, str], list[dict]] = defaultdict(list)
    for r in rows:
        groups[group_key(r)].append(r)
    out: list[dict] = []
    for key, items in groups.items():
        items.sort(key=lambda x: float(x.get("plat_price") or 0), reverse=True)
        gids = [str(i.get("goods_id") or "") for i in items]
        gkey = f"{key[0][:40]}|{key[1]}"
        for it in items:
            it = dict(it)
            it["重复组键"] = gkey
            it["组内商品数"] = len(items)
            it["组内goods_ids"] = ",".join(gids)
            out.append(it)
    out.sort(
        key=lambda x: (
            str(x.get("大类") or ""),
            str(x.get("子类") or ""),
            -float(x.get("plat_price") or 0),
            str(x.get("spu_name") or ""),
        )
    )
    return out


def run(
    input_xlsx: Path,
    profile: Path,
    max_items_per_brand: int | None = None,
    brands_limit: int | None = None,
) -> Path:
    tasks = load_tasks(input_xlsx)
    log(f"总表任务 {len(tasks)} 行")
    brand_to_meta: dict[str, list[dict]] = defaultdict(list)
    for t in tasks:
        brand_to_meta[t["品牌名称"]].append(t)
    brands = list(brand_to_meta.keys())
    if brands_limit:
        brands = brands[:brands_limit]
    log(f"去重品牌 {len(brands)} 个")

    JSONL_DIR.mkdir(parents=True, exist_ok=True)
    all_rows: list[dict] = []
    seen_ids: set[str] = set()
    t0 = time.time()

    with DouhuoClient(profile, headless=True) as client:
        # 等待页面完全就绪，降低价格未加载风险
        client.open_entry()
        client._page.wait_for_timeout(3500)
        log(f"登录态: {client.login_state()}")

        for bi, brand in enumerate(brands, 1):
            metas = brand_to_meta[brand]
            majors = sorted({m["大类"] for m in metas})
            subs = [m["子类"] for m in metas]
            log(f"[{bi}/{len(brands)}] 品牌「{brand}」 大类={majors} 子类数={len(subs)}")
            write_status(f"[品牌品类] {bi}/{len(brands)} {brand} …")
            try:
                goods, total = search_brand_goods(client, brand, max_items=max_items_per_brand)
            except NotLoggedIn:
                log("登录失效，请重新 --login")
                raise
            except Exception as exc:  # noqa: BLE001
                log(f"  品牌「{brand}」失败: {exc}")
                continue

            # 一个品牌可归属多个大类/子类：优先按标题匹配子类；无法匹配则写入该品牌全部大类行
            for item in goods:
                gid = str(item.get("goods_id") or "")
                if gid in seen_ids:
                    continue
                seen_ids.add(gid)
                matched_sub = match_subcategory(str(item.get("spu_name") or ""), brand, subs)
                target_metas = [m for m in metas if (not matched_sub or m["子类"] == matched_sub)]
                if not target_metas:
                    target_metas = metas
                # 若命中多个大类，按第一个 meta 写一行即可（品牌维度主归属）；同时记录匹配子类
                meta0 = target_metas[0]
                row = dict(item)
                row["品牌名称"] = brand
                row["大类"] = meta0["大类"]
                row["子类"] = meta0["子类"]
                row["匹配子类"] = matched_sub
                row["商品链接"] = (
                    f"https://www.douhuomall.com/choice/v3/#/goods?id={gid}&nav_type=1"
                )
                row["采集时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                all_rows.append(row)

            log(
                f"  「{brand}」接口total={total} 新增 {len([r for r in all_rows if r.get('品牌名称')==brand])}"
            )
            # 定时进度
            now = time.time()
            if bi % 5 == 0 or bi == len(brands):
                pct = bi * 100.0 / max(len(brands), 1)
                eta = (now - t0) / bi * (len(brands) - bi) / 60 if bi else 0
                msg = f"[进度] 品牌 {bi}/{len(brands)} ({pct:.0f}%) 累计商品 {len(all_rows)} ETA≈{eta:.0f}分钟"
                log(msg)
                write_status(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

    rows = postprocess(all_rows)
    by_major: dict[str, list[dict]] = defaultdict(list)
    for r in rows:
        by_major[r.get("大类") or "未分类"].append(r)
    log(f"合计商品 {len(rows)} 条，大类 {list(by_major)}")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = OUT_DIR / f"品牌品类选品_{stamp}.xlsx"
    write_workbook(out_path, by_major, VALUE_LEGEND)
    # jsonl 备份
    (JSONL_DIR / f"all_{stamp}.json").write_text(
        json.dumps(rows, ensure_ascii=False), encoding="utf-8"
    )
    log(f"已导出 {out_path}")
    for major, items in by_major.items():
        log(f"  表「{major}」: {len(items)} 行")
    return out_path


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    p = argparse.ArgumentParser(description="品牌品类总表批量抓取")
    p.add_argument("--input", default=str(DEFAULT_INPUT))
    p.add_argument("--profile", default=str(DEFAULT_PROFILE))
    p.add_argument("--max-per-brand", type=int, default=0)
    p.add_argument("--brands-limit", type=int, default=0, help="调试：只跑前 N 个品牌")
    args = p.parse_args()
    try:
        run(
            Path(args.input),
            Path(args.profile),
            max_items_per_brand=args.max_per_brand or None,
            brands_limit=args.brands_limit or None,
        )
    except NotLoggedIn:
        log("请先登录: python douhuo_scraper/run.py --login")
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
