"""
core/native_rip.py — Native COM 视觉渲染封装 (V3.2 Stdin/Stdout IPC 隔离版)
============================================================
通过 Stdin/Stdout 进行纯文本对象交互。
1. 隔离 MCP 线程环境中的 CoInitialize 冲突
2. 避免生成大量临时 JSON 文件导致磁盘碎片
"""

import os
import json
import subprocess
import sys
import tempfile

def render_range_to_png(excel_path: str, sheet_name: str,
                        range_str: str, output_path: str) -> bool:
    """
    通过子进程调用桥接器执行渲染。
    """
    bridge_script = os.path.join(os.path.dirname(__file__), "com_bridge.py")
    
    payload = {
        "cmd": "render",
        "excel_path": excel_path,
        "sheet_name": sheet_name,
        "range_str": range_str,
        "output_path": output_path
    }
    
    try:
        result = subprocess.run(
            [sys.executable, bridge_script],
            input=json.dumps(payload),
            capture_output=True,
            text=True,
            encoding='utf-8'
        )
        
        if result.returncode == 0:
            data = json.loads(result.stdout)
            return data.get("success", False)
        return False
    except Exception:
        return False
