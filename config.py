"""
MCP Server V3 — 全局配置
========================
所有模块共享的常量、路径和审计关键词。
"""

import os

# 默认模板路径
DEFAULT_TEMPLATE = os.path.join(os.path.dirname(__file__), "resources", "template.xlsx")

# 视觉审计缓存目录
CACHE_DIR = os.path.join(os.getcwd(), "audit_cache")

# 高危列关键词（混合简繁中英）
AUDIT_KEYWORDS = [
    "备注", "remarks", "附件", "attachments",
    "项目", "description", "对应单号",
    "audit", "审计", "備注", "note", "detail",
]

# OWASP 公式注入 — 危险前缀
DANGEROUS_PREFIXES = (
    '=', '+', '-', '@', '\t', '\r', '\n',
    '＝', '＋', '－', '＠',
)
