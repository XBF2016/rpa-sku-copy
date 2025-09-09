from pathlib import Path
from typing import List
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from io import BytesIO
import time
try:
    # Pillow 可选依赖，用于将网络图片嵌入到 Excel
    from PIL import Image as PILImage
    _HAS_PILLOW = True
except Exception:
    PILImage = None  # type: ignore
    _HAS_PILLOW = False

# 导出相关工具

def export_results_to_excel(results: List[List[str]], headers: List[str], file_path: Path) -> None:
    """使用 openpyxl 将结果写入 Excel（含表头）。
    行为说明：
    - 若表头包含“图片”列，则尝试下载并嵌入图片；失败则保留为图片链接文本。
    - 若最后一列表头为“价格”，将该列写为纯数值（自动去货币符号与千分位）。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "结果"

    # 写入表头
    ws.append(headers)

    # 识别“图片”与“图片链接”列索引（没有则返回 -1）
    img_col_idx = -1
    img_link_col_idx = -1
    try:
        img_col_idx = headers.index("图片")
    except ValueError:
        img_col_idx = -1
    try:
        img_link_col_idx = headers.index("图片链接")
    except ValueError:
        img_link_col_idx = -1

    # 统一设置图片与链接列列宽（若存在）
    if img_col_idx != -1:
        col_letter = get_column_letter(img_col_idx + 1)
        # 加宽列以便更大预览
        ws.column_dimensions[col_letter].width = 60
    if img_link_col_idx != -1:
        col_letter_link = get_column_letter(img_link_col_idx + 1)
        # 链接列更宽，方便直接查看与复制
        ws.column_dimensions[col_letter_link].width = 90

    for row in results:
        # 价格列数值化（要求“价格”为最后一列表头）
        if headers and headers[-1] == "价格" and any(ch.isdigit() for ch in str(row[-1])):
            row[-1] = float(next((t for t in ''.join((c if (c.isdigit() or c in '.,') else ' ') for c in str(row[-1])).split() if t), '0').replace(',', ''))

        # 先写一整行文本（包括图片URL占位）
        ws.append(row)
        current_row = ws.max_row

        # 若存在图片列，则尝试在该单元格嵌入图片
        if img_col_idx != -1:
            try:
                img_url = str(row[img_col_idx]).strip() if img_col_idx < len(row) else ""
                if _HAS_PILLOW and img_url.startswith("http"):
                    # 按需导入 requests，避免无该依赖时报错
                    try:
                        import requests  # type: ignore
                    except Exception:
                        requests = None  # type: ignore
                    if requests is None:
                        # 无法下载则保留文本链接
                        continue
                    # 下载图片用于嵌入（保留原始分辨率）
                    resp = requests.get(img_url, timeout=10)
                    resp.raise_for_status()
                    bio = BytesIO(resp.content)
                    pil_img = PILImage.open(bio)
                    # 转换成RGB，避免某些模式导致嵌入失败
                    if pil_img.mode not in ("RGB", "RGBA"):
                        pil_img = pil_img.convert("RGB")
                    # 使用单元格坐标锚定（OneCellAnchor）：随单元格移动但不调整大小，保持图片比例
                    xl_img = XLImage(pil_img)
                    anchor = ws.cell(row=current_row, column=img_col_idx + 1).coordinate
                    xl_img.anchor = anchor
                    ws.add_image(xl_img)
                    # 行高设置为较小值，方便整体阅读；图片不会随单元格缩放，比例不变
                    ws.row_dimensions[current_row].height = 180
                # 若无 Pillow 或下载失败，保留文本链接即可
            except Exception:
                # 忽略嵌入失败，仍保留URL文本
                pass

    # 保存工作簿；若被占用则自动降级另存为 result_YYYYMMDD_HHMMSS.xlsx
    try:
        wb.save(str(file_path))
    except Exception as e:
        try:
            ts = time.strftime("%Y%m%d_%H%M%S")
            alt = file_path.with_name(f"{file_path.stem}_{ts}{file_path.suffix}")
            wb.save(str(alt))
            print(f"[警告] Excel 文件被占用或写入失败，已改为另存为: {alt}")
        except Exception as e2:
            print(f"[错误] 导出Excel失败: {e2}")
            raise
