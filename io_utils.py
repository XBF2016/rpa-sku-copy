from pathlib import Path
from typing import List
from openpyxl import Workbook

# 导出相关工具

def export_results_to_excel(results: List[List[str]], headers: List[str], file_path: Path) -> None:
    """使用 openpyxl 将结果写入 Excel（含表头），文件路径由调用方提供。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "结果"
    ws.append(headers)
    for row in results:
        ws.append(row)
    wb.save(str(file_path))
