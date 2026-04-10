"""
core/com_bridge.py — Excel COM 物理隔离桥接器 (v3.2 Stdin/Stdout IPC 版)
==============================================================
作为一个独立的进程运行，通过 Stdin 接收 JSON 指令并通过 Stdout 输出结果。
彻底解决安全风险、临时文件碎片以及 MCP 线程上下文冲突。
"""

import sys
import os
import json
import time
import pythoncom
import win32com.client
from PIL import ImageGrab

def render_range(payload):
    """执行截图渲染逻辑"""
    excel_path = payload["excel_path"]
    sheet_name = payload["sheet_name"]
    range_str = payload["range_str"]
    output_path = payload["output_path"]
    
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(excel_path), ReadOnly=True)
        ws = wb.Sheets(sheet_name)
        ws.Activate()
        
        target_range = ws.Range(range_str)
        target_range.Select()
        
        success = False
        for _ in range(3):
            try:
                target_range.CopyPicture(Appearance=1, Format=2)
                time.sleep(0.5)
                img = ImageGrab.grabclipboard()
                if img:
                    img.save(output_path, "PNG")
                    success = True
                    break
            except:
                time.sleep(1)
        
        wb.Close(SaveChanges=False)
        excel.Quit()
        return {"success": success}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        pythoncom.CoUninitialize()

def inject_padding(payload):
    """执行美学行高注入逻辑"""
    file_path = payload["file_path"]
    sheet_name = payload["sheet_name"]
    start_row = payload["start_row"]
    end_row = payload["end_row"]
    padding_pt = payload["padding_pt"]
    min_height = payload["min_height"]
    max_height = payload["max_height"]
    
    pythoncom.CoInitialize()
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        ws = wb.Sheets(sheet_name)
        ws.Activate()
        
        data_range = ws.Range(f"A{start_row}:U{end_row}")
        data_range.Rows.AutoFit()
        time.sleep(0.3)
        
        clamped = []
        for r in range(start_row, end_row + 1):
            raw_h = ws.Rows(r).RowHeight
            new_h = max(min_height, min(max_height, raw_h + padding_pt))
            ws.Rows(r).RowHeight = new_h
            if new_h >= max_height: clamped.append(r)
            
        wb.Save()
        wb.Close(SaveChanges=True)
        excel.Quit()
        return {"success": True, "clamped": clamped}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    try:
        # Read the full JSON payload from standard input
        input_data = sys.stdin.read()
        if not input_data.strip():
            print(json.dumps({"success": False, "error": "Empty stdin payload"}))
            sys.exit(1)
            
        payload = json.loads(input_data)
            
        cmd = payload.get("cmd")
        if cmd == "render":
            res = render_range(payload)
        elif cmd == "padding":
            res = inject_padding(payload)
        else:
            res = {"success": False, "error": "Unknown command"}
            
        print(json.dumps(res))
    except Exception as e:
        print(json.dumps({"success": False, "error": str(e)}))
