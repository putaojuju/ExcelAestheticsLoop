# MCP Excel Writer (V3.2) 🚀

![V3.2](https://img.shields.io/badge/Architecture-V3.2_Isolated_COM-blue)
![License](https://img.shields.io/badge/License-MIT-green)

An industrial-grade Model Context Protocol (MCP) server for high-precision Excel auditing and layout automation.

## 🌟 Features

- **Three-Tier Isolation Architecture**: Uses a subprocess bridge with Stdin/Stdout JSON streaming to eliminate Windows COM threading conflicts (`CoInitialize` errors) even in high-concurrency async environments.
- **Native Visual RIP**: High-resolution screenshot capture of specific Excel ranges for visual auditing (Source of Truth verification).
- **Aesthetic Layout Engine**: Automated adaptive row-height adjustment (Aesthetics Padding) and quantitative column-width management.
- **Safety First**: OWASP-inspired value sanitization and mandatory file backups before every write operation.

## 🏗️ Architecture

The system is split into three layers to ensure stability:
1. **Routing Layer (FastMCP)**: Asynchronous MCP dispatcher (No win32com imports).
2. **IPC Layer (Stream JSON)**: Stdin/Stdout JSON streaming between processes.
3. **Execution Layer (Isolated Bridge)**: Independent Python process for handling sensitive COM automation.

## 📦 Installation

```bash
pip install fastmcp openpyxl pywin32 Pillow
```

## 🚀 Configuration

Add the following to your MCP settings (`claude_desktop_config.json` or `mcp_config.json`):

```json
{
  "mcpServers": {
    "excel-writer": {
      "command": "python",
      "args": ["/path/to/excel_mcp_v3.py"],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

## 🛡️ License

MIT License. Developed for the Siu Hong (兆康) Settlement Project.
