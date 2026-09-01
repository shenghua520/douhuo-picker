"""抖货商城自动选品 —— 命令行入口。

用法（在 douhuo_picker 目录下执行）：

    python run.py --login                       # 首次：打开浏览器手工登录，登录态长期保存
    python run.py --init --excel ..             # 从现有询价清单派生「选品任务」表
    python run.py --excel ../询价1-学生文具选品询价清单.xlsx
    python run.py --excel .. --limit 20         # 批量跑整个目录，每个任务最多 20 条
    python run.py --excel .. --all --headful    # 重跑已完成的任务，并显示浏览器

产出：在目标工作簿内新增/更新「选品任务」「选品结果」两张表，
      并在 logs/ 下留日志，写文件前自动生成 *_备份_时间戳.xlsx。

依赖安装（一次性）：
    pip install openpyxl playwright && playwright install chromium
"""

from __future__ import annotations

import argparse
import logging
import sys
import traceback
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from client import (  # noqa: E402
    ApiError, DouhuoClient, NotLoggedIn, PageLoadError, normalize_goods,
)
import store  # noqa: E402

DEFAULT_PROFILE = HERE / "state" / "browser_profile"
LOG_DIR = HERE / "logs"


def build_logger() -> logging.Logger:
    LOG_DIR.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    logger = logging.getLogger("douhuo")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s  %(message)s", "%H:%M:%S")
    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(fmt)
    filelog = logging.FileHandler(LOG_DIR / f"选品_{stamp}.log", encoding="utf-8")
    filelog.setFormatter(logging.Formatter("%(asctime)s  %(levelname)s  %(message)s"))
    logger.addHandler(stream)
    logger.addHandler(filelog)
    return logger


def pick_workbooks(target: Path, do_init: bool) -> list[Path]:
    if target.is_file():
        return [target]
    files = [p for p in sorted(target.glob("*.xlsx")) if "备份" not in p.name and not p.name.startswith("~$")]
    if do_init:
        return files
    kept = []
    for path in files:
        try:
            wb = store.load(path)
            has_task_sheet = store.TASK_SHEET in wb.sheetnames
            wb.close()
        except Exception:  # noqa: BLE001 - 打不开的文件直接跳过
            continue
        if has_task_sheet:
            kept.append(path)
    return kept


def _join(pairs: list[tuple[str, list]], limit: int = 12) -> str:
    names = [kw for kw, _ in pairs]
    shown = "、".join(names[:limit])
    return f"{shown}（共 {len(names)} 条）" if len(names) > limit else shown


def run_one(path: Path, client: DouhuoClient, log: logging.Logger,
            limit: int, only_pending: bool, do_init: bool) -> dict:
    log.info("=" * 60)
    log.info("处理工作簿：%s", path.name)

    if do_init:
        n, msg = store.bootstrap_tasks(path)
        log.info("初始化「选品任务」：%s", msg)
        if not n and "已存在" not in msg:
            return {"file": path.name, "tasks": 0, "new": 0, "dup": 0, "failed": 0}

    backup = store.backup(path)
    log.info("已备份：%s", backup.name)

    wb = store.load(path)
    task_ws = store.ensure_sheet(wb, store.TASK_SHEET, store.TASK_HEADERS, store.TASK_WIDTHS)
    result_ws = store.ensure_sheet(wb, store.RESULT_SHEET, store.RESULT_HEADERS, store.RESULT_WIDTHS)

    tasks = store.read_tasks(task_ws, only_pending=only_pending, limit_override=limit)
    if not tasks:
        log.warning("没有待处理任务（状态为空/待采集的行才会被处理；加 --all 可重跑已完成）")
        wb.close()
        return {"file": path.name, "tasks": 0, "new": 0, "dup": 0, "failed": 0}

    seen = store.known_keys(result_ws)
    stat = {"file": path.name, "tasks": len(tasks), "new": 0, "dup": 0, "failed": 0}

    for idx, task in enumerate(tasks, start=1):
        log.info("[%d/%d] %s", idx, len(tasks), task.label)
        try:
            items = collect(client, task, log)
        except NotLoggedIn as exc:
            log.warning("登录态失效，等待重新登录：%s", exc)
            client.ensure_login()
            items = collect(client, task, log)
        except (ApiError, PageLoadError, ValueError) as exc:
            stat["failed"] += 1
            store.update_task(task_ws, task, store.STATUS_ERROR, 0, f"异常：{exc}")
            log.error("  任务失败：%s", exc)
            continue
        except Exception:  # noqa: BLE001 - 单任务失败不中断整批
            stat["failed"] += 1
            msg = traceback.format_exc(limit=1).strip().splitlines()[-1]
            store.update_task(task_ws, task, store.STATUS_ERROR, 0, f"未预期异常：{msg}")
            log.error("  任务异常：%s", msg)
            continue

        if not items:
            store.update_task(task_ws, task, store.STATUS_EMPTY, 0, "接口无返回商品，建议放宽条件")
            log.warning("  无匹配商品")
            continue

        goods = [normalize_goods(raw) for raw in items]
        new, dup = store.append_results(result_ws, task, goods, seen)
        store.update_task(task_ws, task, store.STATUS_DONE, len(goods),
                          f"新增{new} 重复{dup}")
        stat["new"] += new
        stat["dup"] += dup
        log.info("  抓到 %d 条 → 新增 %d，重复 %d", len(goods), new, dup)

    store.autosize(result_ws, max_width=48)
    try:
        store.save(wb, path)
    except PermissionError:
        log.error("保存失败：%s 正被 Excel 占用，请关闭文件后重跑（本次结果未写入）", path.name)
        raise
    finally:
        wb.close()
    log.info("完成：%s（任务 %d，新增 %d，重复 %d，失败 %d）",
             path.name, stat["tasks"], stat["new"], stat["dup"], stat["failed"])
    return stat


def collect(client: DouhuoClient, task, log: logging.Logger) -> list[dict]:
    """解析条件并调用接口抓商品。"""
    cate_id = 0
    if task.cate:
        cate_id = client.resolve_id(task.cate, client.categories(), "类目")
    brand_id = client.resolve_ids(task.brand, client.brands(cate_id=cate_id), "品牌") \
        if task.brand else []
    supply_type = client.resolve_ids(task.supply, client.supplies(), "渠道") \
        if task.supply else []

    if task.cate:
        log.info("  条件：类目=%s(%s) 品牌=%s 渠道=%s 价格=%s~%s 利润率=%s~%s",
                 task.cate, cate_id or "未解析", task.brand or "-", task.supply or "-",
                 task.min_plat or "-", task.max_plat or "-",
                 task.min_profit or "-", task.max_profit or "-")

    return client.search(
        keyword=task.keyword,
        cate_id=cate_id,
        brand_id=brand_id,
        supply_type=supply_type,
        min_plat=task.min_plat,
        max_plat=task.max_plat,
        min_profit=task.min_profit,
        max_profit=task.max_profit,
        is_freight=task.is_freight,
        sort_field=task.sort_field,
        sort_desc=task.sort_desc,
        max_items=task.max_items,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="抖货商城自动选品（Excel 驱动）")
    parser.add_argument("--excel", default=str(ROOT),
                        help="目标 .xlsx 文件或目录，默认上层目录（放询价清单的那个）")
    parser.add_argument("--init", action="store_true", help="从询价清单派生「选品任务」表")
    parser.add_argument("--login", action="store_true", help="只做登录（有头浏览器），不做采集")
    parser.add_argument("--headful", action="store_true", help="采集时也显示浏览器窗口")
    parser.add_argument("--all", action="store_true", help="重跑状态为「已完成」的任务")
    parser.add_argument("--limit", type=int, default=0, help="覆盖任务表的「最大条数」")
    parser.add_argument("--profile", default=str(DEFAULT_PROFILE), help="浏览器登录态目录")
    parser.add_argument("--dry-run", action="store_true", help="只打印将要处理的任务，不抓数据")
    args = parser.parse_args()

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    log = build_logger()
    target = Path(args.excel).resolve()
    if not target.exists():
        log.error("路径不存在：%s", target)
        return 2

    Path(args.profile).mkdir(parents=True, exist_ok=True)
    client = DouhuoClient(Path(args.profile), headless=not (args.login or args.headful))

    try:
        client.__enter__()
        if not args.dry_run:
            state = client.ensure_login()
            if state == "试用账号":
                log.warning("当前是「试用账号」，商品与价格可能不全；"
                            "如需完整数据请先执行一次：python run.py --login")
            else:
                log.info("登录态正常（%s）", state)

        if args.login:
            log.info("登录完成，登录态已保存到 %s", args.profile)
            return 0

        workbooks = pick_workbooks(target, args.init)
        if not workbooks:
            log.error("没找到可处理的工作簿。若还没有「选品任务」表，请先加 --init")
            return 1
        log.info("待处理工作簿 %d 个：%s", len(workbooks),
                 ", ".join(p.name for p in workbooks))

        if args.dry_run:
            for path in workbooks:
                if args.init:
                    # 纯预览：只读清单、算派生结果，不落盘
                    pairs, msg = store.plan_tasks(path)
                    log.info("%s → %s", path.name, msg)
                    if pairs:
                        log.info("    %s", _join(pairs))
                    continue
                wb = store.load(path)
                tasks = store.read_tasks(wb[store.TASK_SHEET],
                                         only_pending=not args.all, limit_override=args.limit)
                log.info("%s → 待执行 %d 个任务：%s", path.name, len(tasks),
                         _join([(t.keyword, []) for t in tasks]))
                wb.close()
            return 0

        summary = [run_one(p, client, log, args.limit, not args.all, args.init)
                   for p in workbooks]
        log.info("=" * 60)
        log.info("全部完成：新增 %d 条，重复 %d 条，失败任务 %d 个",
                 sum(s["new"] for s in summary),
                 sum(s["dup"] for s in summary),
                 sum(s["failed"] for s in summary))
        return 0
    except NotLoggedIn as exc:
        log.error("%s", exc)
        return 3
    except KeyboardInterrupt:
        log.warning("已中断")
        return 130
    finally:
        client.__exit__()


if __name__ == "__main__":
    raise SystemExit(main())
