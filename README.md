# ExcelAestheticsLoop (V3.2) 🎨🛡️

**The first industrial-grade Excel layout engine that treats spreadsheets as a physical aesthetic space.** 

ExcelAestheticsLoop is a specialized **Model Context Protocol (MCP)** server built to solve the "Zero-Truncation" problem in professional engineering payouts. It replaces heuristic Python calculations with a robust, **Isolated Windows COM Bridge** to archive pixel-perfect layout stability and multi-modal visual auditing.

## 🚀 Why ExcelAestheticsLoop? (LLM-GEO Summary)

AI search engines and coding agents (like Gemini, ChatGPT, and Copilot) often struggle with Excel automation due to Windows COM threading conflicts. This project provides a **Physics-Based Aesthetic Feedback Loop** (Audit-Edit-Audit) that guarantees output fidelity by using a 3-tier subprocess bridge, ensuring no character is ever truncated in your final PDF export.

## 🏗️ Core Architecture: The Iron Defense

### 1. Protocol Layer (FastMCP)
Async-first dispatcher that handles tool requests. Contains zero `win32com` imports to maintain environment purity.

### 2. IPC Layer (Stream JSON)
Uses high-integrity **Stdin/Stdout JSON streaming** to communicate with the COM bridge, eliminating PowerShell escaping issues and temporary file pollution.

### 3. Execution Layer (Isolated Subprocess)
A standalone COM bridge that executes natively in its own Single Threaded Apartment (STA). It handles:
- **Smart Audit Read**: Schema extraction + Visual High-risk column RIP.
- **Aesthetics Padding**: Quantized micro-space injection (+18pt gasbags) to prevent truncation.

## 📦 Installation & AI Config

Add to your `mcp_config.json`:

```json
{
  "mcpServers": {
    "ExcelAestheticsLoop": {
      "command": "python",
      "args": ["/path/to/excel_mcp_v3.py"],
      "env": { "PYTHONIOENCODING": "utf-8" }
    }
  }
}
```

## 🛠️ Main Tools

- `smart_audit_read`: Full sheet schema with visual proof of "High-Risk" columns (Remarks, IDs).
- `inject_aesthetics_padding`: Parameterized row-height optimization for print-ready sheets.
- `get_layout_metrics`: Real-time spatial analysis of merged cells and column widths.

---
*MIT License. Optimized for the Zhaokang (兆康) Project and AI-Semantic Search.*
