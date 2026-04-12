# Excel - MCP Server (V3.2)

基于 Model Context Protocol (MCP) 的 Excel 自动化工具。解决 LLM 调用 Excel 时的稳定性问题，支持自动调整行高、防止内容截断。

## 核心功能

- **行高调整**：可供模型自主调整行高，确保导出无截断。
- **稳定性增强**：通过子进程隔离技术，解决 Windows 环境下的 COM 组件冲突。
- **视觉审计**：支持对 Excel 指定区域进行截图。

## 安装与配置

在 MCP 配置文件（如 `claude_desktop_config.json`）中添加：

```json
{
  "mcpServers": {
    "excel-mcp": {
      "command": "python",
      "args": ["E:/lin/.agents/mcp_server/excel_mcp_v3.py"],
      "env": { "PYTHONIOENCODING": "utf-8" }
    }
  }
}
```

## 工具列表

- `smart_audit_read`: 读取表格结构并截图。
- `inject_aesthetics_padding`: 可供模型自主调整行高，防止打印截断。
- `get_layout_metrics`: 分析合并单元格和列宽参数。

---
*MIT License.*
