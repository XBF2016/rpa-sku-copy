from pathlib import Path
from typing import List
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, TwoCellAnchor
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
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

def _col_width_to_px(col_width: float) -> float:
    """将 Excel 列宽(字符数)近似换算为像素。
    经验公式（Calibri 11 下基本适用）：px = col_width * 7 + 5
    """
    try:
        return float(col_width) * 7.0 + 5.0
    except Exception:
        return 8.43 * 7.0 + 5.0  # 默认列宽对应像素


def _px_to_col_width(px: float) -> float:
    """将像素近似换算为 Excel 列宽(字符数)。"""
    return max(1.0, (float(px) - 5.0) / 7.0)


def _px_to_points(px: float) -> float:
    """像素 -> point（1 pt = 1/72 in, 96 DPI 假定）。"""
    return float(px) * 72.0 / 96.0

def export_results_to_excel(results: List[List[str]], headers: List[str], file_path: Path) -> None:
    """使用 openpyxl 将结果写入 Excel（含表头）。
    行为说明：
    - 若表头包含“图片”列，则下载并按【图片列列宽】等比例缩放后锚定到该单元格内，尽量避免溢出；失败则保留为文本链接。
    - 全表单元格统一设置为：上下左右居中、自动换行；
    - 文本列列宽依据内容长度进行近似自适应；行高对含图行按图片高度自适应，其余行交由 Excel 自动。
    - 若最后一列表头为“价格”，将该列写为纯数值（自动去货币符号与千分位）。
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "结果"

    # 写入表头
    ws.append(headers)
    # 表头样式：居中 + 自动换行
    for c in ws[1]:
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 统计每列最大文本长度（用于后续列宽自适应）
    max_text_len = [len(str(h)) for h in headers]

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
        # 使用较小的缩略图列宽，图片会按该列宽进行等比缩放后嵌入到单元格
        _IMG_PREVIEW_PX = 160  # 目标缩略图宽度（像素）
        ws.column_dimensions[col_letter].width = _px_to_col_width(_IMG_PREVIEW_PX)
    if img_link_col_idx != -1:
        col_letter_link = get_column_letter(img_link_col_idx + 1)
        # 链接列更宽，方便直接查看与复制
        ws.column_dimensions[col_letter_link].width = 90

    for row in results:
        # 价格列数值化（要求“价格”为最后一列表头）
        if headers and headers[-1] == "价格" and any(ch.isdigit() for ch in str(row[-1])):
            row[-1] = float(next((t for t in ''.join((c if (c.isdigit() or c in '.,') else ' ') for c in str(row[-1])).split() if t), '0').replace(',', ''))

        # 先写一整行文本；若存在“图片”列，则该列不写入文本，只留空用于放图
        row_for_write = list(row)
        if img_col_idx != -1 and img_col_idx < len(row_for_write):
            row_for_write[img_col_idx] = ""
        ws.append(row_for_write)
        current_row = ws.max_row

        # 当前行单元格统一样式（上下左右居中、自动换行）
        for ci in range(1, len(headers) + 1):
            ws.cell(row=current_row, column=ci).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # 统计文本长度用于自适应列宽（图片列跳过）
        for j in range(len(headers)):
            if j == img_col_idx:
                continue
            try:
                txt_len = len(str(row[j])) if j < len(row) else 0
            except Exception:
                txt_len = 0
            if j < len(max_text_len) and txt_len > max_text_len[j]:
                max_text_len[j] = txt_len

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
                    if requests is not None:
                        # 下载图片用于嵌入
                        resp = requests.get(img_url, timeout=10)
                        resp.raise_for_status()
                        bio = BytesIO(resp.content)
                        pil_img = PILImage.open(bio)
                        # 转换成RGB，避免某些模式导致嵌入失败
                        if pil_img.mode not in ("RGB", "RGBA"):
                            pil_img = pil_img.convert("RGB")
                        xl_img = XLImage(pil_img)

                        # 依据图片列列宽按比例缩放，力求完全“装入”单元格
                        img_col_letter = get_column_letter(img_col_idx + 1)
                        cur_cw = ws.column_dimensions[img_col_letter].width or 8.43
                        col_px = _col_width_to_px(cur_cw)
                        # 预留少量内边距，避免紧贴边框
                        target_w_px = max(1, int(col_px - 8))
                        scale = min(1.0, target_w_px / max(1, pil_img.width))
                        new_w_px = max(1, int(pil_img.width * scale))
                        new_h_px = max(1, int(pil_img.height * scale))
                        xl_img.width = new_w_px
                        xl_img.height = new_h_px

                        # 使用 TwoCellAnchor：左上角定位到该单元格，右下角在同一单元格内按像素偏移，达到“贴边且随单元格缩放”的效果
                        tl_col = img_col_idx
                        tl_row = current_row - 1
                        from_marker = AnchorMarker(col=tl_col, colOff=pixels_to_EMU(0), row=tl_row, rowOff=pixels_to_EMU(0))
                        to_marker = AnchorMarker(col=tl_col, colOff=pixels_to_EMU(new_w_px), row=tl_row, rowOff=pixels_to_EMU(new_h_px))
                        xl_img.anchor = TwoCellAnchor(_from=from_marker, to=to_marker, editAs="twoCell")
                        ws.add_image(xl_img)

                        # 行高与图片高度一致（point），避免图片溢出到相邻行
                        target_row_h_pts = _px_to_points(new_h_px + 4)
                        cur_h = ws.row_dimensions[current_row].height or 0
                        ws.row_dimensions[current_row].height = max(cur_h, target_row_h_pts)
                # 若无 Pillow 或下载失败，保留文本链接即可
            except Exception:
                # 忽略嵌入失败，仍保留URL文本
                pass

    # 文本列近似自适应列宽（图片列与图片链接列除外）
    for j in range(len(headers)):
        if j == img_col_idx or j == img_link_col_idx:
            continue
        col_letter_j = get_column_letter(j + 1)
        # 控制在 [10, 40] 以内，避免过窄或过宽
        target_width = max(10, min(40, (max_text_len[j] if j < len(max_text_len) else 10) + 2))
        ws.column_dimensions[col_letter_j].width = target_width

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
