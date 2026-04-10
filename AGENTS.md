# AGENTS.md - Instruction Guide for Autonomous Agents

This repository is designed for high-precision spreadsheet manipulation. Autonomous agents (Cursor, Copilot, etc.) should follow these rules to maintain environment stability.

## Environment Initialization
- **Runtime**: Windows (Required for COM).
- **Python**: 3.10+
- **Dependencies**: `pip install fastmcp openpyxl pywin32 Pillow`

## Architecture Constraints (V3.2)
- **NO Direct COM Integration**: Never import `win32com.client` in any service file. This causes thread context deadlocks within the FastMCP environment.
- **Subprocess Bridge Protocol**: All COM operations must be routed through `core/com_bridge.py` using the Stdin/Stdout JSON protocol to ensure apartment thread safety.
- **Data Integrity**: Every write operation must verify that a `.bak` file was created by `io_service.py` prior to finalizing changes.

## Testing Protocol
1. **Connectivity Check**: Execute `get_layout_metrics` on a sample file to verify the subprocess bridge integrity before attempting complex row-height injections.
2. **IPC Validation**: If communication failures occur, verify that `PYTHONIOENCODING=utf-8` is set in the execution environment.

## Coding Standards
- Implement strict Type Hinting (Python 3.10+ PEP 585/604).
- Use `pathlib.Path` for all filesystem operations.
- Document implementation logic in Docstrings to facilitate semantic RAG retrieval by AI engines.
