"""选品中心五表全量抓取入口。

用法：
  python run.py --login
  python run.py --scrape all
  python run.py --scrape banner,daily,tuan_top,tuan_all,youpin
  python run.py --export-only
"""

from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from client import API, DouhuoClient, NotLoggedIn  # noqa: E402
from excelio import (  # noqa: E402
    BANNER_PREFERRED,
    GOODS_PREFERRED,
    normalize_banner_row,
    normalize_goods_row,
    write_excel,
)

DEFAULT_PROFILE = ROOT / "state" / "browser_profile"
DATA_DIR = ROOT / "output" / "jsonl"
XLSX_DIR = ROOT / "output"

SECTIONS = ["banner", "daily", "tuan_top", "tuan_all", "youpin"]

SECTION_FILES = {
    "banner": "01_首页轮播图商品.xlsx",
    "daily": "02_首页每日必看专区商品.xlsx",
    "tuan_top": "03_团购爆品轮播图商品.xlsx",
    "tuan_all": "04_团购爆品所有商品.xlsx",
    "youpin": "05_海量优品全部商品.xlsx",
}


def log(msg: str) -> None:
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}", flush=True)


def jsonl_path(section: str) -> Path:
    return DATA_DIR / f"{section}.jsonl"


def load_jsonl(section: str) -> list[dict]:
    path = jsonl_path(section)
    if not path.exists():
        return []
    rows = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError:
                continue
    return rows


def append_jsonl(section: str, rows: list[dict]) -> None:
    if not rows:
        return
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with jsonl_path(section).open("a", encoding="utf-8") as f:
        for row in rows:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")


def reset_jsonl(section: str) -> None:
    path = jsonl_path(section)
    if path.exists():
        path.unlink()


def goods_key(row: dict) -> str:
    return str(row.get("goods_id") or row.get("id") or "")


def _enrich_goods_detail(client: DouhuoClient, goods_id: Any) -> dict:
    """尽量用 getGoodsDetail 补全字段；失败则返回空 dict。"""
    if goods_id in ("", None):
        return {}
    try:
        detail = client.post(API["goods_detail"], {"goods_id": goods_id}) or {}
    except Exception:  # noqa: BLE001
        return {}
    if not isinstance(detail, dict):
        return {}
    # detail_img 是数组，写 Excel 时会 json 化；这里先压成字符串减负
    if isinstance(detail.get("detail_img"), list):
        detail["detail_img"] = ";".join(str(x) for x in detail["detail_img"] if x)
    return detail


def scrape_banner(client: DouhuoClient, fresh: bool) -> list[dict]:
    """首页轮播区：
    1) banner 图元数据
    2) 右侧 sright 商品（getTopSpecial.goods_list）
    3) 每张 banner link_url 跳转专区/商品的全量列表
    """
    if fresh:
        reset_jsonl("banner")
    rows: list[dict] = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    data = client.banner()
    banners = data.get("banner_img") or []
    base = data.get("base_map") or {}
    if isinstance(base, dict) and base:
        banners = list(banners) + [base]
    banner_rows = []
    for b in banners:
        parsed = client.parse_link_ids(str(b.get("link_url") or b.get("mobile_url") or ""))
        row = normalize_banner_row(b, parsed)
        row["_row_type"] = "banner_img"
        banner_rows.append(row)
    rows.extend(banner_rows)
    log(f"首页轮播图 banner 图: {len(banner_rows)} 条")

    specials = client.top_special()
    goods_rows: list[dict] = []
    for sp in specials:
        name = sp.get("nav_name") or ""
        nav_id = sp.get("nav_id")
        meta = {k: v for k, v in sp.items() if k != "goods_list"}
        meta["_row_type"] = "special_meta"
        meta["nav_id"] = nav_id
        meta["nav_name"] = name
        meta["goods_count"] = len(sp.get("goods_list") or [])
        meta["采集时间"] = now
        rows.append(meta)
        for g in sp.get("goods_list") or []:
            mapped = {
                "goods_id": g.get("goods_id"),
                "spu_name": g.get("goods_name") or g.get("spu_name"),
                "main_img": g.get("cover_img") or g.get("main_img"),
                "plat_price": g.get("plat_price"),
                "retail_price": g.get("retail_price"),
                "profit": g.get("profit"),
            }
            for k, v in g.items():
                if k not in mapped:
                    mapped[k] = v
            detail = _enrich_goods_detail(client, mapped.get("goods_id"))
            for k, v in detail.items():
                if k not in mapped or mapped.get(k) in ("", None):
                    mapped[k] = v
            if detail.get("goods_name") and not mapped.get("spu_name"):
                mapped["spu_name"] = detail.get("goods_name")
            if detail.get("cover_img") and not mapped.get("main_img"):
                mapped["main_img"] = detail.get("cover_img")
            goods_rows.append(
                normalize_goods_row(
                    mapped, nav_id=nav_id, nav_name=name, source="首页轮播图sright"
                )
            )
    log(f"首页轮播侧栏商品: {len(goods_rows)} 条（来自 getTopSpecial）")

    # 轮播图跳转链接里的商品
    link_goods = 0
    for brow in banner_rows:
        banner_id = brow.get("img_id")
        banner_name = brow.get("img_name") or f"banner_{banner_id}"
        link_url = brow.get("link_url") or brow.get("mobile_url") or ""
        nav_id = brow.get("解析nav_id") or ""
        nav_type = int(brow.get("解析nav_type") or 1 or 0)
        goods_id = brow.get("解析goods_id") or ""
        tags = {
            "来源banner_id": banner_id,
            "来源banner名称": banner_name,
            "来源banner链接": link_url,
        }

        # 单品跳转
        if goods_id:
            try:
                res = client.goods_list(page=1, limit=20, **{"type": 1, "spu_name": "", "spu_sn": ""})
                # 更稳：用详情接口
                detail = _enrich_goods_detail(client, goods_id)
                item = {"goods_id": goods_id, **{k: v for k, v in detail.items() if v not in ("", None)}}
                if not detail:
                    item = {"goods_id": goods_id}
                row = normalize_goods_row(
                    item, nav_id="", nav_name=banner_name, source="轮播图跳转单品"
                )
                row.update(tags)
                rows.append(row)
                link_goods += 1
                log(f"  轮播「{banner_name}」单品 goods_id={goods_id} 已收录")
            except Exception as exc:  # noqa: BLE001
                log(f"  轮播「{banner_name}」单品失败: {exc}")
            continue

        if not nav_id:
            continue

        log(f"  开始采集轮播「{banner_name}」跳转专区 nav_id={nav_id} …")
        got = 0
        batch: list[dict] = []
        try:
            page = 1
            while True:
                res = client.goods_list(nav_type=nav_type, nav_id=nav_id, page=page, limit=100)
                items = res.get("list") or []
                pages = res.get("pages")
                if not items:
                    break
                for item in items:
                    row = normalize_goods_row(
                        item, nav_id=nav_id, nav_name=banner_name, source="轮播图跳转专区"
                    )
                    row.update(tags)
                    batch.append(row)
                    got += 1
                if len(batch) >= 100:
                    rows.extend(batch)
                    batch = []
                if pages is not None and page >= int(pages):
                    break
                page += 1
                if page % 5 == 0:
                    log(f"    「{banner_name}」 page={page} 已取 {got} 条…")
            if batch:
                rows.extend(batch)
            link_goods += got
            log(f"  轮播「{banner_name}」跳转专区共 {got} 条商品")
        except Exception as exc:  # noqa: BLE001
            if batch:
                rows.extend(batch)
            if "1014" in str(exc) or "权限" in str(exc):
                log(f"  轮播「{banner_name}」专区无权限，已跳过: {exc}")
            else:
                log(f"  轮播「{banner_name}」专区失败: {exc}")

    log(f"轮播图跳转链接商品合计: {link_goods} 条")
    rows.extend(goods_rows)
    append_jsonl("banner", rows)
    return load_jsonl("banner")


def scrape_daily(client: DouhuoClient, fresh: bool, max_per_nav: int | None = None) -> list[dict]:
    if fresh:
        reset_jsonl("daily")
    existing = load_jsonl("daily")
    done_goods = {goods_key(r) for r in existing if goods_key(r)}
    navs = client.recs_special()
    log(f"每日必看入口: {len(navs)} 个专区")
    total_new = 0
    # 海量优品/应有尽有等与全库同量级的专栏不放进「每日必看」，避免重复抓 10 万+
    skip_navs = {7, 8}
    for idx, nav in enumerate(navs, 1):
        nav_id = nav.get("nav_id")
        nav_name = nav.get("nav_name") or ""
        if nav_id in skip_navs:
            log(f"  [{idx}/{len(navs)}] 跳过全库专栏 {nav_name}({nav_id})（由海量优品表覆盖）")
            continue
        # 入口元数据单独一行（无 goods_id）
        meta = dict(nav)
        meta["_row_type"] = "nav_meta"
        meta["nav_id"] = nav_id
        meta["nav_name"] = nav_name
        meta["采集时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        append_jsonl("daily", [meta])
        batch: list[dict] = []
        try:
            for item in client.iter_goods(
                nav_type=int(nav.get("nav_type") or 1),
                nav_id=nav_id,
                limit=100,
                max_items=max_per_nav,
            ):
                gid = goods_key(item)
                if gid in done_goods:
                    continue
                done_goods.add(gid)
                batch.append(
                    normalize_goods_row(item, nav_id=nav_id, nav_name=nav_name, source="每日必看")
                )
                if len(batch) >= 100:
                    append_jsonl("daily", batch)
                    total_new += len(batch)
                    batch = []
            if batch:
                append_jsonl("daily", batch)
                total_new += len(batch)
            log(f"  [{idx}/{len(navs)}] {nav_name}({nav_id}) 完成")
        except Exception as exc:  # noqa: BLE001
            if "1014" in str(exc) or "权限不足" in str(exc):
                log(f"  [{idx}/{len(navs)}] 跳过无权限专栏 {nav_name}({nav_id}): {exc}")
            else:
                log(f"  [{idx}/{len(navs)}] {nav_name}({nav_id}) 失败: {exc}")
            if batch:
                append_jsonl("daily", batch)
                total_new += len(batch)
    log(f"每日必看商品新增约 {total_new} 条")
    return load_jsonl("daily")


def scrape_tuan_top(client: DouhuoClient, fresh: bool) -> list[dict]:
    if fresh:
        reset_jsonl("tuan_top")
    specials = client.top_special()
    rows = []
    for sp in specials:
        name = sp.get("nav_name") or ""
        nav_id = sp.get("nav_id")
        goods = sp.get("goods_list") or []
        # 入口元信息
        meta = {k: v for k, v in sp.items() if k != "goods_list"}
        meta["_row_type"] = "special_meta"
        meta["goods_count"] = len(goods)
        meta["采集时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        rows.append(meta)
        if "团购" in name or str(nav_id) == "11":
            for g in goods:
                # 轮播字段较少：goods_id/goods_name/cover_img/plat_price/retail_price/profit
                mapped = {
                    "goods_id": g.get("goods_id"),
                    "spu_name": g.get("goods_name") or g.get("spu_name"),
                    "main_img": g.get("cover_img") or g.get("main_img"),
                    "plat_price": g.get("plat_price"),
                    "retail_price": g.get("retail_price"),
                    "profit": g.get("profit"),
                }
                mapped.update({k: v for k, v in g.items() if k not in mapped})
                rows.append(
                    normalize_goods_row(
                        mapped, nav_id=nav_id, nav_name=name, source="团购爆品轮播"
                    )
                )
    append_jsonl("tuan_top", rows)
    log(f"团购爆品轮播: 落盘 {len(rows)} 行（含入口元数据）")
    return load_jsonl("tuan_top")


def progress_path(section: str) -> Path:
    return DATA_DIR / f"{section}.progress.json"


def read_progress(section: str) -> dict:
    path = progress_path(section)
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


STATUS_FILE = ROOT / "output" / "progress_status.txt"


def write_progress(section: str, page: int, collected: int, total: object = None) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    progress_path(section).write_text(
        json.dumps(
            {"next_page": page, "collected": collected, "total": total, "updated": time.time()},
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )


def write_status(text: str) -> None:
    """给人看的进度文件，便于另开窗口 tail。"""
    try:
        STATUS_FILE.parent.mkdir(parents=True, exist_ok=True)
        STATUS_FILE.write_text(text + "\n", encoding="utf-8")
    except OSError:
        pass


def scrape_paged_goods(
    client: DouhuoClient,
    section: str,
    *,
    nav_type: int,
    nav_id: str | int,
    source: str,
    nav_name: str,
    fresh: bool,
    max_items: int | None = None,
) -> list[dict]:
    if fresh:
        reset_jsonl(section)
        p = progress_path(section)
        if p.exists():
            p.unlink()
    existing = load_jsonl(section)
    done = {goods_key(r) for r in existing if goods_key(r)}
    collected = len(done)
    start_page = int(read_progress(section).get("next_page") or 1)
    log(
        f"{section}: 已有 {collected} 条，从第 {start_page} 页续抓 nav_id={nav_id} "
        f"(limit=100, max={max_items or 'ALL'})"
    )
    write_status(f"[{section}] 开始 从第{start_page}页 已有{collected}条")
    batch: list[dict] = []
    newly = 0
    page = start_page
    t0 = time.time()
    last_log = t0
    log_every = 200  # 条
    try:
        while True:
            res = client.goods_list(nav_type=nav_type, nav_id=nav_id, page=page, limit=100)
            items = res.get("list") or []
            total = res.get("total")
            pages = res.get("pages")
            if not items:
                write_progress(section, page, collected + newly, total)
                break
            for item in items:
                gid = goods_key(item)
                if gid in done:
                    continue
                done.add(gid)
                batch.append(
                    normalize_goods_row(item, nav_id=nav_id, nav_name=nav_name, source=source)
                )
                newly += 1
                if max_items is not None and newly >= max_items:
                    if batch:
                        append_jsonl(section, batch)
                    write_progress(section, page + 1, collected + newly, total)
                    log(f"{section} 达到 max_items={max_items}，累计 {collected + newly}")
                    return load_jsonl(section)
            if batch:
                append_jsonl(section, batch)
                batch = []
            page += 1
            write_progress(section, page, collected + newly, total)

            now = time.time()
            should_log = (newly and newly % log_every < 100) or (now - last_log >= 20)
            if should_log:
                got = collected + newly
                pct = ""
                eta = ""
                if total:
                    try:
                        total_n = int(total)
                        pct = f"{got * 100.0 / total_n:.1f}%"
                        if got > collected and now > t0:
                            rate = (got - collected) / max(now - t0, 1e-6)
                            remain = max(total_n - got, 0)
                            mins = remain / rate / 60 if rate > 0 else 0
                            eta = f" ETA≈{mins:.0f}分钟"
                    except (TypeError, ValueError):
                        pass
                msg = (
                    f"[进度] {section} 已写 {got}"
                    f"{f' / total={total}' if total else ''} "
                    f"page={page - 1}{f' {pct}' if pct else ''}{eta}"
                )
                log(msg)
                write_status(
                    f"[{datetime.now().strftime('%H:%M:%S')}] {msg} 速度≈"
                    f"{((got - collected) / max(now - t0, 1e-6)):.1f}条/秒"
                )
                last_log = now
            if pages is not None and (page - 1) >= int(pages):
                break
        if batch:
            append_jsonl(section, batch)
    except KeyboardInterrupt:
        if batch:
            append_jsonl(section, batch)
        write_progress(section, page, collected + newly)
        log(f"{section} 中断，已保存 {collected + newly} 条，下次从第 {page} 页续")
        raise
    log(f"{section} 完成，累计 {collected + newly} 条")
    return load_jsonl(section)


def export_section(section: str) -> Path:
    rows = load_jsonl(section)
    out = XLSX_DIR / SECTION_FILES[section]
    if section == "banner":
        n = write_excel(out, rows, BANNER_PREFERRED)
    else:
        # nav_meta rows and goods rows share sheet; goods preferred first
        n = write_excel(out, rows, GOODS_PREFERRED)
    log(f"导出 {out.name}: {n} 行")
    return out


def export_all() -> list[Path]:
    return [export_section(s) for s in SECTIONS]


def scrape(
    client: DouhuoClient,
    sections: list[str],
    *,
    fresh: bool,
    youpin_max: int | None = None,
    daily_max_per_nav: int | None = None,
) -> None:
    state = client.login_state()
    log(f"登录态: {state}")
    if "banner" in sections:
        scrape_banner(client, fresh=fresh)
        export_section("banner")
    if "daily" in sections:
        scrape_daily(client, fresh=fresh, max_per_nav=daily_max_per_nav)
        export_section("daily")
    if "tuan_top" in sections:
        scrape_tuan_top(client, fresh=fresh)
        export_section("tuan_top")
    if "tuan_all" in sections:
        scrape_paged_goods(
            client,
            "tuan_all",
            nav_type=1,
            nav_id=11,
            source="团购爆品全量",
            nav_name="团购爆品",
            fresh=fresh,
        )
        export_section("tuan_all")
    if "youpin" in sections:
        # 优先专栏 nav_id=7；无权限时回退全站
        used_nav = "7"
        try:
            probe = client.goods_list(nav_type=1, nav_id=7, page=1, limit=1)
            log(f"海量优品专栏 total={probe.get('total')}")
        except Exception as exc:  # noqa: BLE001
            log(f"nav_id=7 不可用({exc})，回退全站 getGoodsList")
            used_nav = ""
        scrape_paged_goods(
            client,
            "youpin",
            nav_type=1 if used_nav else 0,
            nav_id=used_nav,
            source="海量优品",
            nav_name="海量优品",
            fresh=fresh,
            max_items=youpin_max,
        )
        export_section("youpin")


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    parser = argparse.ArgumentParser(description="抖货选品中心五表全量抓取")
    parser.add_argument("--login", action="store_true", help="有头登录并保存 profile")
    parser.add_argument(
        "--scrape",
        default="",
        help="all 或逗号分隔: banner,daily,tuan_top,tuan_all,youpin",
    )
    parser.add_argument("--export-only", action="store_true", help="只从 jsonl 导出 Excel")
    parser.add_argument("--fresh", action="store_true", help="清空对应分区 jsonl 后重抓")
    parser.add_argument("--profile", default=str(DEFAULT_PROFILE))
    parser.add_argument("--headful", action="store_true")
    parser.add_argument("--youpin-max", type=int, default=0, help="海量优品最大条数，0=不限")
    parser.add_argument("--daily-max-per-nav", type=int, default=0)
    args = parser.parse_args()

    if args.export_only:
        for p in export_all():
            log(f"OK {p}")
        return 0

    sections: list[str] = []
    if args.scrape:
        sections = SECTIONS if args.scrape.strip() == "all" else [
            s.strip() for s in args.scrape.split(",") if s.strip()
        ]
        unknown = [s for s in sections if s not in SECTIONS]
        if unknown:
            log(f"未知分区: {unknown}")
            return 2

    profile = Path(args.profile)
    profile.mkdir(parents=True, exist_ok=True)
    headless = not (args.login or args.headful)

    with DouhuoClient(profile, headless=headless) as client:
        if args.login:
            state = client.ensure_login()
            log(f"登录完成: {state}")
            return 0
        if not sections:
            log("请指定 --login 或 --scrape all")
            return 2
        try:
            scrape(
                client,
                sections,
                fresh=args.fresh,
                youpin_max=args.youpin_max or None,
                daily_max_per_nav=args.daily_max_per_nav or None,
            )
        except NotLoggedIn as exc:
            log(f"登录态问题: {exc}")
            log("请先执行: python douhuo_scraper/run.py --login")
            return 3
    log("全部完成")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
