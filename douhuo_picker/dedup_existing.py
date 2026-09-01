"""就地校正已生成的 9 份表：把同一任务内「商品ID」或「名称+双价」重复的行标记为「重复」并删除。"""

from __future__ import annotations

import glob
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

import store
from store import RESULT_HEADERS, _C_TASK, _C_GOODS, _C_NAME, _C_PLAT, _C_RETAIL

DUP_FILL = PatternFill("solid", fgColor="FFE0B2")


def dedup_one(path: Path) -> tuple[int, int, int]:
    """返回 (原始行数, 删除行数, 唯一行数)。"""
    wb = openpyxl.load_workbook(path)
    if "选品结果" not in wb.sheetnames:
        wb.close()
        return 0, 0, 0
    ws = wb["选品结果"]
    seen, keep, row = set(), [], 2
    while row <= ws.max_row:
        cell = ws.cell(row=row, column=_C_TASK).value
        t_id = str(cell or "").strip()
        keys = store.dedup_keys(
            t_id,
            ws.cell(row=row, column=_C_GOODS).value,
            ws.cell(row=row, column=_C_NAME).value,
            ws.cell(row=row, column=_C_PLAT).value,
            ws.cell(row=row, column=_C_RETAIL).value,
        )
        if any(k in seen for k in keys):
            ws.delete_rows(row, 1)
            continue
        seen.update(keys)
        keep.append(row)
        row += 1
    n, dup, kept = ws.max_row - 1, ws.max_row - 1 - len(keep), len(keep)
    wb.save(path)
    wb.close()
    return n, dup, kept


def main() -> int:
    files = sorted(glob.glob(str(HERE.parent / "询价*.xlsx")))
    if not files:
        print("未找到询价清单"); return 1
    total_orig = total_dup = 0
    for f in files:
        orig, dup, kept = dedup_one(Path(f))
        total_orig += orig
        total_dup += dup
        if dup:
            print(f"  {Path(f).name:<34} 原 {orig:>4} → 去重 {dup:>3} 行 → 剩 {kept:>4}")
        else:
            print(f"  {Path(f).name:<34} 原 {orig:>4} → 无重复")
    print("-" * 60)
    print(f"合计：原 {total_orig} 行，移除 {total_dup} 条重复，剩 {total_orig - total_dup} 条")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
