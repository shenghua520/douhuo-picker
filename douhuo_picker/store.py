"""Excel 读写层：选品任务表（输入）+ 选品结果表（输出）。

约定：在同一个 .xlsx 工作簿里挂两张表，不改用户原有的询价清单内容。
- 「选品任务」：人填条件，脚本读
- 「选品结果」：脚本写，人查阅/导出

只做读条件 / 写结果 / 去重这三件事，不做模板引擎。
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from client import SORT_DESC, SORT_FIELD

TASK_SHEET = "选品任务"
RESULT_SHEET = "选品结果"

# 输入列：人只需要填这些
TASK_HEADERS = [
    "任务ID", "关键词", "类目", "品牌", "渠道", "价格下限", "价格上限",
    "利润率下限", "利润率上限", "仅包邮", "排序", "最大条数", "状态", "命中数", "备注",
]
# 输出列
RESULT_HEADERS = [
    "任务ID", "关键词", "商品ID", "商品名称", "渠道", "代发价", "零售价", "集采价",
    "利润率(%)", "预估佣金", "包邮", "推荐指数", "上架状态", "商品链接", "主图",
    "采集时间", "状态",
]

# 子表表头/汇总行会被当成数据行，列出来过滤掉
SKIP_KEYWORDS = {"系列", "品牌", "品类", "规格", "序号", "商品名称", "备注", "说明"}

STATUS_PENDING, STATUS_DONE = "待采集", "已完成"
STATUS_EMPTY, STATUS_ERROR = "无匹配", "异常"
MARK_NEW, MARK_DUP = "新增", "重复"

HEADER_FILL = PatternFill("solid", fgColor="1D1F27")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
OK_FILL = PatternFill("solid", fgColor="E8F5E9")
WARN_FILL = PatternFill("solid", fgColor="FFF8E1")
ERR_FILL = PatternFill("solid", fgColor="FDECEA")


@dataclass
class Task:
    """一行选品任务。row 用于回写状态。"""

    row: int
    task_id: str
    keyword: str
    cate: str
    brand: str
    supply: str
    min_plat: float
    max_plat: float
    min_profit: float
    max_profit: float
    is_freight: bool
    sort_field: int
    sort_desc: int
    max_items: int
    note: str = ""

    @property
    def label(self) -> str:
        return f"[{self.task_id}] {self.keyword or '(无关键词)'}"


def backup(path: Path) -> Path:
    """写文件前留一份带时间戳的备份，避免脚本异常损坏原始清单。"""
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst = path.with_name(f"{path.stem}_备份_{stamp}{path.suffix}")
    shutil.copy2(path, dst)
    return dst


def load(path: Path):
    return load_workbook(path)


def save(wb, path: Path) -> None:
    wb.save(path)


# ---------- 建表 ----------

def _style_header(ws: Worksheet, headers: list[str], widths: list[int]) -> None:
    for idx, (name, width) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=1, column=idx, value=name)
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"


def ensure_sheet(wb, name: str, headers: list[str], widths: list[int]) -> Worksheet:
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    _style_header(ws, headers, widths)
    return ws


TASK_WIDTHS = [10, 26, 14, 14, 14, 10, 10, 11, 11, 8, 16, 10, 10, 8, 34]
RESULT_WIDTHS = [10, 22, 12, 44, 14, 10, 10, 14, 11, 11, 8, 10, 10, 52, 40, 20, 8]


def plan_tasks(path: Path) -> tuple[list[tuple[str, list[str]]], str]:
    """只读：从工作簿第一张（询价清单）派生 (关键词, 品牌列表)。不动文件。

    适配两种表头：.../品类/品牌/... 与 .../品牌/系列/...
    """
    wb = load(path)
    try:
        if TASK_SHEET in wb.sheetnames and wb[TASK_SHEET].max_row > 1:
            return [], f"「{TASK_SHEET}」已存在且有数据，跳过生成"

        src = wb.worksheets[0]
        header_row = _find_header_row(src)
        if header_row is None:
            return [], f"未能在「{src.title}」中定位表头（需含『品类』或『品牌』列），请手工填「{TASK_SHEET}」"

        headers = [str(src.cell(row=header_row, column=c).value or "").strip()
                   for c in range(1, src.max_column + 1)]
        kw_col = _pick_col(headers, "品类") or _pick_col(headers, "系列")
        brand_col = _pick_col(headers, "品牌")
        if kw_col is None:
            return [], "未找到『品类』或『系列』列，请手工填「选品任务」"

        # 一个品类在清单里往往有多行（不同品牌/规格），合并成一条任务，
        # 品牌名汇总进备注做参考，不写进「品牌」列——否则会把搜索范围限制死。
        merged: dict[str, list[str]] = {}
        for r in range(header_row + 1, src.max_row + 1):
            keyword = str(src.cell(row=r, column=kw_col).value or "").strip()
            if not keyword or keyword in SKIP_KEYWORDS or keyword.startswith(("小计", "合计", "总计")):
                continue
            brand = str(src.cell(row=r, column=brand_col).value or "").strip() if brand_col else ""
            bucket = merged.setdefault(keyword, [])
            if brand and brand not in SKIP_KEYWORDS and brand not in bucket:
                bucket.append(brand)

        return list(merged.items()), f"可从「{src.title}」派生 {len(merged)} 条任务（按品类合并去重）"
    finally:
        wb.close()


def bootstrap_tasks(path: Path) -> tuple[int, str]:
    """plan_tasks 的结果落盘成「选品任务」表。"""
    pairs, msg = plan_tasks(path)
    if not pairs:
        return 0, msg

    wb = load(path)
    ws = ensure_sheet(wb, TASK_SHEET, TASK_HEADERS, TASK_WIDTHS)
    for count, (keyword, brands) in enumerate(pairs, start=1):
        ws.append([
            f"T{count:03d}", keyword, "", "", "", "", "", "", "", "",
            "", "", 40, STATUS_PENDING, "",
            f"原清单品牌：{'/'.join(brands)}" if brands else "",
        ])
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=13, max_col=13):
        for cell in row:
            cell.alignment = Alignment(horizontal="center")
    save(wb, path)
    wb.close()
    return len(pairs), f"已写入 {len(pairs)} 条任务到「{TASK_SHEET}」"


def _find_header_row(ws: Worksheet) -> int | None:
    limit = min(ws.max_row, 12)
    for r in range(1, limit + 1):
        cells = [str(ws.cell(row=r, column=c).value or "").strip()
                 for c in range(1, min(ws.max_column, 12) + 1)]
        if any("品类" in x or x == "品牌" for x in cells) and sum(bool(x) for x in cells) >= 3:
            return r
    return None


def _pick_col(headers: list[str], key: str) -> int | None:
    for idx, name in enumerate(headers, start=1):
        if key in name:
            return idx
    return None


# ---------- 读任务 ----------

def read_tasks(ws: Worksheet, only_pending: bool = True, limit_override: int = 0) -> list[Task]:
    tasks: list[Task] = []
    for r in range(2, ws.max_row + 1):
        values = [ws.cell(row=r, column=c).value for c in range(1, len(TASK_HEADERS) + 1)]
        row = dict(zip(TASK_HEADERS, values))
        keyword = str(row.get("关键词") or "").strip()
        cate = str(row.get("类目") or "").strip()
        if not keyword and not cate:
            continue
        status = str(row.get("状态") or "").strip()
        if only_pending and status in (STATUS_DONE,):
            continue
        sort_field, sort_desc = _parse_sort(str(row.get("排序") or ""))
        tasks.append(Task(
            row=r,
            task_id=str(row.get("任务ID") or f"R{r}").strip(),
            keyword=keyword,
            cate=cate,
            brand=str(row.get("品牌") or "").strip(),
            supply=str(row.get("渠道") or "").strip(),
            min_plat=_num(row.get("价格下限")),
            max_plat=_num(row.get("价格上限")),
            min_profit=_num(row.get("利润率下限")),
            max_profit=_num(row.get("利润率上限")),
            is_freight=str(row.get("仅包邮") or "").strip() in ("1", "是", "Y", "y", "TRUE", "True"),
            sort_field=sort_field,
            sort_desc=sort_desc,
            max_items=int(limit_override or _num(row.get("最大条数")) or 40),
            note=str(row.get("备注") or "").strip(),
        ))
    return tasks


SORT_LABELS = {"利润率": "profit", "零售价": "retail", "成本价": "cost"}


def _parse_sort(text: str) -> tuple[int, int]:
    """接受『利润率降序 / profit_desc / 零售价升序 / 成本价降序』等写法。"""
    text = str(text or "").strip().lower()
    if not text:
        return 0, 0
    order = "asc" if ("升" in text or "asc" in text) else "desc"
    for label, key in SORT_LABELS.items():
        if label in text or key in text:
            return SORT_FIELD[key], SORT_DESC[order]
    return 0, 0


def _num(value) -> float:
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return 0.0


# ---------- 写结果 ----------

_C_TASK, _C_GOODS = RESULT_HEADERS.index("任务ID") + 1, RESULT_HEADERS.index("商品ID") + 1
_C_NAME, _C_PLAT = RESULT_HEADERS.index("商品名称") + 1, RESULT_HEADERS.index("代发价") + 1
_C_RETAIL = RESULT_HEADERS.index("零售价") + 1


def dedup_keys(task_id: str, goods_id, name, plat, retail) -> tuple:
    """同任务内双键判重：商品ID + 「名称+双价」。

    平台存在同一商品重复挂 listing（goods_id 不同但名称与价格完全一样）的情况，
    只按 goods_id 判重会漏掉它们；只按名称+双价判重又会因同名不同规格被错杀，
    所以双键并用。
    """
    return ((task_id, str(goods_id or "").strip()),
            (task_id, str(name or "").strip(), plat, retail))


def known_keys(ws: Worksheet) -> set:
    """已采集过的结果行判重键集合，用于同任务重跑时判重。"""
    keys = set()
    for r in range(2, ws.max_row + 1):
        cell = lambda c: ws.cell(row=r, column=c).value  # noqa: E731
        keys.update(dedup_keys(
            str(cell(_C_TASK) or "").strip(), cell(_C_GOODS),
            cell(_C_NAME), cell(_C_PLAT), cell(_C_RETAIL),
        ))
    return keys


def append_results(ws: Worksheet, task: Task, goods: list[dict], seen
                   ) -> tuple[int, int]:
    """写入一批商品，返回 (新增数, 重复数)。seen 是双键集合（商品ID + 名称+双价）。"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new = dup = 0
    for item in goods:
        keys = dedup_keys(task.task_id, item["goods_id"], item["spu_name"],
                          item["plat_price"], item["retail_price"])
        if any(k in seen for k in keys):
            mark = MARK_DUP
            dup += 1
        else:
            mark = MARK_NEW
            seen.update(keys)
            new += 1
        row = [
            task.task_id, task.keyword, item["goods_id"], item["spu_name"], item["supply_name"],
            item["plat_price"], item["retail_price"], item["purchase_price"], item["profit"],
            item["commission"], item["is_freight"], item["sort"], item["status"],
            item["url"], item["main_img"], now, mark,
        ]
        ws.append(row)
        cell = ws.cell(row=ws.max_row, column=len(RESULT_HEADERS))
        cell.fill = OK_FILL if mark == MARK_NEW else WARN_FILL
        if item["url"]:
            link = ws.cell(row=ws.max_row, column=RESULT_HEADERS.index("商品链接") + 1)
            link.hyperlink = item["url"]
            link.style = "Hyperlink"
    return new, dup


def update_task(ws: Worksheet, task: Task, status: str, hits: int, note: str = "") -> None:
    ws.cell(row=task.row, column=TASK_HEADERS.index("状态") + 1, value=status)
    ws.cell(row=task.row, column=TASK_HEADERS.index("命中数") + 1, value=hits)
    if note:
        ws.cell(row=task.row, column=TASK_HEADERS.index("备注") + 1, value=note[:250])
    cell = ws.cell(row=task.row, column=TASK_HEADERS.index("状态") + 1)
    cell.fill = {STATUS_DONE: OK_FILL, STATUS_EMPTY: WARN_FILL}.get(status, ERR_FILL)
    cell.alignment = Alignment(horizontal="center")


def autosize(ws: Worksheet, max_width: int = 60) -> None:
    for column in ws.columns:
        letter = get_column_letter(column[0].column)
        longest = max((len(str(c.value or "")) for c in column), default=8)
        ws.column_dimensions[letter].width = min(max(longest + 2, 10), max_width)
