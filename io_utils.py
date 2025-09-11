from pathlib import Path
from typing import List
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import time
import hashlib
from urllib.parse import urlparse
try:
    import requests  # 可选依赖，用于下载图片到本地
except Exception:
    requests = None  # type: ignore
try:
    from PIL import Image as PILImage
    _HAS_PILLOW = True
except Exception:
    PILImage = None  # type: ignore
    _HAS_PILLOW = False
try:
    import win32com.client as win32  # 可选依赖，用于通过 Excel COM 插入“链接的图片”
except Exception:
    win32 = None  # type: ignore

# 导出相关工具

def _px_to_col_width(px: float) -> float:
    """将像素近似换算为 Excel 列宽(字符数)（保留，供可能的复用）。"""
    try:
        return max(1.0, (float(px) - 5.0) / 7.0)
    except Exception:
        return 12.0

def _px_to_points(px: float) -> float:
    """像素 -> point（1 pt = 1/72 in, 96 DPI 假定）。"""
    return float(px) * 72.0 / 96.0


def _insert_linked_images_via_excel(xlsx_path: Path, img_col_idx: int, url_to_filename: dict, results: List[List[str]]) -> None:
    """使用 Excel COM（win32com）在目标列插入“链接的图片”，不将图片数据保存进 Excel。
    - LinkToFile=True, SaveWithDocument=False
    - 形状随单元格移动/缩放，按单元格尺寸等比缩放并居中。
    """
    if img_col_idx == -1 or not results:
        return
    if win32 is None:
        try:
            print("[提示] 未安装 win32com（pywin32），将不会在表格中插入可见图片，仅保留链接列。")
        except Exception:
            pass
        return

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    try:
        wb = excel.Workbooks.Open(str(xlsx_path))
        ws = wb.Worksheets(1)

        # 删除目标列（从第2行开始）已有的形状，避免重复叠加
        try:
            for i in range(ws.Shapes.Count, 0, -1):
                shp = ws.Shapes.Item(i)
                try:
                    if shp.TopLeftCell.Column == (img_col_idx + 1) and shp.TopLeftCell.Row >= 2:
                        shp.Delete()
                except Exception:
                    pass
        except Exception:
            pass

        # 逐行插入“链接的图片”
        for r_idx, row in enumerate(results, start=2):  # Excel 行号从2开始（第1行为表头）
            try:
                url = str(row[img_col_idx]).strip()
            except Exception:
                url = ""
            fname = url_to_filename.get(url)
            if not fname:
                continue
            png_path = (xlsx_path.parent / fname).resolve()
            if not png_path.exists():
                continue

            cell = ws.Cells(r_idx, img_col_idx + 1)
            left = float(cell.Left)
            top = float(cell.Top)
            try:
                shp = ws.Shapes.AddPicture(str(png_path), True, False, left, top, -1, -1)
            except Exception:
                # 某些版本可能需要 AddPicture2，但通常 AddPicture 即可
                continue

            try:
                # 随单元格移动与缩放
                shp.Placement = 1  # xlMoveAndSize
                shp.LockAspectRatio = True
                # 将链接源改为相对文件名，便于与工作簿同目录移动
                try:
                    shp.LinkFormat.SourceFullName = png_path.name
                except Exception:
                    pass
                try:
                    shp.AlternativeText = png_path.name
                except Exception:
                    pass
                # 按单元格完整尺寸等比缩放（无内边距）
                cell_w = float(cell.Width)
                cell_h = float(cell.Height)
                if cell_w > 0 and cell_h > 0 and shp.Width and shp.Height:
                    scale = min(cell_w / max(1.0, float(shp.Width)), cell_h / max(1.0, float(shp.Height)), 1.0)
                    shp.Width = max(1.0, float(shp.Width) * scale)
                    shp.Height = max(1.0, float(shp.Height) * scale)
                # 居中对齐（至少有一边会完全贴合单元格边界，比例不变）
                shp.Left = float(cell.Left) + max(0.0, (float(cell.Width) - float(shp.Width)) / 2.0)
                shp.Top = float(cell.Top) + max(0.0, (float(cell.Height) - float(shp.Height)) / 2.0)
            except Exception:
                pass

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        try:
            excel.Quit()
        except Exception:
            pass

def export_results_to_excel(results: List[List[str]], headers: List[str], file_path: Path) -> None:
    """使用 openpyxl 将结果写入 Excel（含表头）。
    行为说明：
    - 若表头包含“图片”列：对图片链接去重后，仅下载一次并统一转为 PNG 保存到与 result.xlsx 同级目录；然后在“图片”列中按列宽嵌入显示（可见缩略图）。
      - 无 Pillow 或下载/转换失败时，退化为在“图片”列写入原始 URL 的超链接。
    - “图片链接”列保留原始网络 URL。
    - 全表统一样式：水平/垂直居中，自动换行；文本列近似自适应列宽。
    - 若最后一列表头为“价格”，将该列写为纯数值（去掉人民币符号与千分位）。
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

    # 去重图片链接并下载到与 Excel 同级目录（统一转为 PNG）
    img_output_dir = file_path.parent
    url_to_filename = {}
    if img_col_idx != -1:
        unique_urls = []
        seen = set()
        for row in results:
            try:
                u = str(row[img_col_idx]).strip()
            except Exception:
                u = ""
            if u and u.startswith("http") and u not in seen:
                seen.add(u)
                unique_urls.append(u)

        downloaded = 0
        for u in unique_urls:
            short = hashlib.md5(u.encode("utf-8")).hexdigest()[:10]
            fname = f"img_{short}.png"  # 统一转为 PNG
            target = img_output_dir / fname
            # 若 requests 或 Pillow 不可用，则直接标记为 None：后续回退为 URL
            if (requests is None) or (not _HAS_PILLOW):
                url_to_filename[u] = None
                continue
            url_to_filename[u] = fname  # 先建立映射（即便下载/转换失败也可回退）
            if target.exists():
                continue
            try:
                resp = requests.get(u, timeout=15)
                resp.raise_for_status()
                from io import BytesIO
                bio = BytesIO(resp.content)
                pil_img = PILImage.open(bio)
                # 规整模式，避免某些色彩模式导致保存失败
                if pil_img.mode not in ("RGB", "RGBA"):
                    pil_img = pil_img.convert("RGB")
                # 统一转为 PNG 保存到磁盘
                pil_img.save(target, format="PNG")
                downloaded += 1
            except Exception:
                url_to_filename[u] = None  # 下载失败：后续回退为URL
        try:
            if unique_urls:
                print(f"[步骤] 检测到 {len(unique_urls)} 个唯一图片链接，已下载 {downloaded} 张到: {img_output_dir}")
        except Exception:
            pass

    # 统一设置图片与链接列列宽（若存在）
    if img_col_idx != -1:
        col_letter = get_column_letter(img_col_idx + 1)
        # 设定预览列宽（像素 -> 列宽），适当加大到 240px
        _IMG_PREVIEW_PX = 240
        ws.column_dimensions[col_letter].width = _px_to_col_width(_IMG_PREVIEW_PX)
    if img_link_col_idx != -1:
        col_letter_link = get_column_letter(img_link_col_idx + 1)
        # “图片链接”列更宽，便于查看与复制
        ws.column_dimensions[col_letter_link].width = 90

    for row in results:
        # 价格列数值化（要求“价格”为最后一列表头）
        if headers and headers[-1] == "价格" and any(ch.isdigit() for ch in str(row[-1])):
            row[-1] = float(next((t for t in ''.join((c if (c.isdigit() or c in '.,') else ' ') for c in str(row[-1])).split() if t), '0').replace(',', ''))

        # 先写一整行文本；若存在“图片”列，先留空，后续在单元格中嵌入缩略图
        row_for_write = list(row)
        local_name = None
        img_url_cur = None
        if img_col_idx != -1 and img_col_idx < len(row_for_write):
            try:
                img_url_cur = str(row[img_col_idx]).strip()
            except Exception:
                img_url_cur = ""
            local_name = url_to_filename.get(img_url_cur) if 'url_to_filename' in locals() else None  # type: ignore
            row_for_write[img_col_idx] = ""
        ws.append(row_for_write)
        current_row = ws.max_row

        # 当前行单元格统一样式（上下左右居中、自动换行）
        for ci in range(1, len(headers) + 1):
            ws.cell(row=current_row, column=ci).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # 统计文本长度用于自适应列宽（以写入后的内容计算，图片列跳过）
        for j in range(len(headers)):
            try:
                if j == img_col_idx:
                    continue
                txt = row_for_write[j] if j < len(row_for_write) else ""
                txt_len = len(str(txt))
            except Exception:
                txt_len = 0
            if j < len(max_text_len) and txt_len > max_text_len[j]:
                max_text_len[j] = txt_len

        # 不再由 openpyxl 嵌入图片，这里不做处理；显示交由 Excel COM 完成

    # 文本列近似自适应列宽（跳过“图片”与“图片链接”列）
    for j in range(len(headers)):
        if j == img_col_idx or j == img_link_col_idx:
            continue
        col_letter_j = get_column_letter(j + 1)
        # 控制在 [10, 40] 以内，避免过窄或过宽
        target_width = max(10, min(40, (max_text_len[j] if j < len(max_text_len) else 10) + 2))
        ws.column_dimensions[col_letter_j].width = target_width

    # 若存在图片列，给数据行一个适中的行高，便于后续 COM 插入的图片按单元格自适应可见
    if img_col_idx != -1:
        try:
            # 行高与预览保持一致（像素 -> point），此处取 240px，保证图片更清晰
            _IMG_PREVIEW_PX_H = 240
            target_row_h_pts = _px_to_points(_IMG_PREVIEW_PX_H)
            for r in range(2, ws.max_row + 1):
                cur_h = ws.row_dimensions[r].height or 0
                ws.row_dimensions[r].height = max(cur_h, target_row_h_pts)
        except Exception:
            pass

    # 保存工作簿；若被占用则自动降级另存为 result_YYYYMMDD_HHMMSS.xlsx
    final_path = file_path
    try:
        wb.save(str(final_path))
    except Exception as e:
        try:
            ts = time.strftime("%Y%m%d_%H%M%S")
            alt = final_path.with_name(f"{final_path.stem}_{ts}{final_path.suffix}")
            wb.save(str(alt))
            final_path = alt
            print(f"[警告] Excel 文件被占用或写入失败，已改为另存为: {alt}")
        except Exception as e2:
            print(f"[错误] 导出Excel失败: {e2}")
            raise

    # 使用 Excel COM 将图片以“链接方式”插入到单元格中（不占用 Excel 体积）
    try:
        _insert_linked_images_via_excel(final_path, img_col_idx, url_to_filename, results)
    except Exception:
        # 插入失败不影响导出
        pass
