from pathlib import Path
from typing import List
from openpyxl import Workbook

# 导出相关工具

def export_results_to_excel(results: List[List[str]], headers: List[str], file_path: Path) -> None:
    """使用 openpyxl 将结果写入 Excel（含表头）。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "结果"

    # 写入表头
    ws.append(headers)

    for row in results:
        # 若最后一列为“价格”，将其转为纯数值（取第一个数值，兼容区间与千分位；无数字则保持原样）
        if headers and headers[-1] == "价格" and any(ch.isdigit() for ch in str(row[-1])):
            row[-1] = float(next((t for t in ''.join((c if (c.isdigit() or c in '.,') else ' ') for c in str(row[-1])).split() if t), '0').replace(',', ''))
        ws.append(row)

    wb.save(str(file_path))
