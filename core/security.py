"""
core/security.py — OWASP 公式注入防御 + 文件备份
=================================================
"""

import os
import shutil
from datetime import datetime
from config import DANGEROUS_PREFIXES


def sanitize_cell_value(value):
    """
    OWASP 标准：阻断公式注入。
    数值类型直接通过；字符串类型检查危险前缀，
    发现则加前置单引号强制文本模式。
    """
    if value is None:
        return value
    if isinstance(value, (int, float)):
        return value
    s = str(value)
    if s and s[0] in DANGEROUS_PREFIXES:
        return "'" + s
    return value


def make_backup(filepath: str) -> str:
    """写操作前自动备份，返回备份路径。"""
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base, ext = os.path.splitext(filepath)
    bak_path = f"{base}.{ts}.bak{ext}"
    shutil.copy2(filepath, bak_path)
    return bak_path
