# AGENTS.md - Instruction Guide for Autonomous Agents

This repository is designed for high-precision spreadsheet manipulation. Autonomous agents (Cursor, Copilot, etc.) should follow these rules to maintain environment stability.

## 🤖 Environment Initialization
- **Runtime**: Windows (Required for COM).
- **Python**: 3.10+
- **Dependencies**: `pip install fastmcp openpyxl pywin32 Pillow`

## 🏗️ Architecture constraints (V3.2)
- **NO Direct COM**: Never import `win32com.client` in any service file. This causes thread context deadlocks.
- **Subprocess Bridge**: All COM operations MUST be routed through `core/com_bridge.py` using the Stdin/Stdout JSON protocol.
- **Backups**: Every write operation must verify that a `.bak` file was created by `io_service.py`.

## 🧪 Testing Protocol
1. **Tool Check**: Test `get_layout_metrics` on a sample file before attempting complex row-height injections.
2. **IPC Validation**: If the bridge fails, check `PYTHONIOENCODING=utf-8` in your environment.

## 📌 Coding Style
- Use Typing Hints (Python 3.10+ PEP 585/604).
- Prefer `pathlib.Path` over os.path strings.
- Implementation details should be documented in Docstrings for better RAG retrieval.
