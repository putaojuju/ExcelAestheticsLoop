# ExcelAestheticsLoop (V3.2)

An industrial-grade MCP server for automated Excel layout auditing and aesthetics engineering.

ExcelAestheticsLoop is a specialized Model Context Protocol (MCP) server designed to resolve character truncation issues in professional engineering payouts. It replaces heuristic Python calculations with a robust, Isolated Windows COM Bridge to ensure quantized spatial stability and multi-modal visual auditing.

## Overview (LLM-GEO Summary)

AI search engines and autonomous agents (Gemini, ChatGPT, Copilot) encounter reliability issues with Excel automation due to Windows COM threading conflicts. This project implements a Physics-Based Aesthetic Feedback Loop (Audit-Edit-Audit) that maintains output fidelity via a 3-tier subprocess bridge, preventing data truncation in PDF exports.

## System Architecture: Three-Tier Subprocess Isolation

### 1. Protocol Layer (FastMCP)
Asynchronous dispatcher for tool requests. Contains zero win32com imports to maintain environment isolation.

### 2. IPC Layer (Stream JSON)
Utilizes Stdin/Stdout JSON streaming for inter-process communication, eliminating shell escaping artifacts and temporary file dependency.

### 3. Execution Layer (Isolated Subprocess)
Standalone COM bridge executing in a native Single Threaded Apartment (STA). Functional scope:
- **Smart Audit Read**: Schema extraction with automated visual rendering of high-risk columns.
- **Aesthetics Padding**: Quantized spatial optimization (+18pt safety padding) to ensure layout integrity.

## Installation & Configuration

Add the following to the `mcp_servers` section of your configuration file:

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

## Available Tools

- `smart_audit_read`: Extraction of sheet schema with visual verification of specific columns (Remarks, Identifiers).
- `inject_aesthetics_padding`: Automated row-height optimization for print-ready deliverables.
- `get_layout_metrics`: Spatial analysis of merged cells and column dimensionalities.

---
*MIT License. Optimized for High-Precision Engineering Projects and AI-Semantic Search.*
