"""
Excel MCP Server V3 — 统一入口
================================
模块化架构：所有业务逻辑分散在 services/ 下，
本文件仅负责初始化 FastMCP 并注册工具路由。

依赖：pip install fastmcp openpyxl pywin32 Pillow
"""

import io
import sys
import os

# 强制 UTF-8 输出
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 让 absolute imports 正确解析
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from fastmcp import FastMCP

mcp = FastMCP("Excel MCP Server V3")

# ────────────────────────────────────────
# 注册所有服务模块的工具
# ────────────────────────────────────────
from services.io_service import register_io_tools
from services.vision_service import register_vision_tools
from services.layout_service import register_layout_tools

register_io_tools(mcp)
register_vision_tools(mcp)
register_layout_tools(mcp)

# ────────────────────────────────────────
if __name__ == "__main__":
    mcp.run()
