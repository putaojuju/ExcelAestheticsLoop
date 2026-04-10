"""
core/excel_mapping.py — 列号转换 + MergedCell 智能路由
=====================================================
"""

from openpyxl.utils import get_column_letter


def col_letter_to_index(letter: str) -> int:
    """将列字母 (A, B, ..., AA) 转为 1-indexed 数字。"""
    letter = letter.upper().strip()
    result = 0
    for ch in letter:
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def resolve_col(col) -> int:
    """接受列字母或数字，统一返回 1-indexed 列号。"""
    if isinstance(col, int):
        return col
    if isinstance(col, str):
        if col.isdigit():
            return int(col)
        return col_letter_to_index(col)
    return int(col)


def int_to_col_letter(n: int) -> str:
    """1-indexed 数字转列字母 (1 -> A)。"""
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def get_primary_cell(ws, row: int, col_letter: str) -> str:
    """
    MergedCell 智能路由：
    如果目标坐标在合并区域内，返回左上角主单元格的坐标。
    否则原样返回。
    """
    coord = f"{col_letter}{row}"
    for merged_range in ws.merged_cells.ranges:
        if coord in merged_range:
            return merged_range.start_cell.coordinate
    return coord
