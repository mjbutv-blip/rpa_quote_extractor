import os
import tempfile
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils.units import pixels_to_EMU
from PIL import Image as PILImage

# 单元格内用于显示图片的最大可用区域（像素）
BOX_W_PX = 220
BOX_H_PX = 220


def _scale_to_fit(src_w: int, src_h: int, box_w: int, box_h: int) -> tuple:
    """取宽缩放比与高缩放比中的较小值，保证图片完整放入且不变形。"""
    scale = min(box_w / src_w, box_h / src_h)
    return max(1, round(src_w * scale)), max(1, round(src_h * scale))


def _resize_image(src_path: str, box_w: int = BOX_W_PX, box_h: int = BOX_H_PX) -> tuple:
    """按等比例缩放算法，将图片缩放至 box 内最大尺寸（不拉伸变形）。

    返回 (新临时文件路径, 缩放后宽px, 缩放后高px)。
    """
    img = PILImage.open(src_path).convert("RGBA")
    target_w, target_h = _scale_to_fit(img.width, img.height, box_w, box_h)
    img = img.resize((target_w, target_h), PILImage.LANCZOS)
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".png")
    os.close(tmp_fd)
    img.save(tmp_path, format="PNG", optimize=False)
    return tmp_path, target_w, target_h


def _px_to_row_height(px: int) -> float:
    """像素转 Excel 行高（点单位，1pt ≈ 1.333px @ 96dpi）。"""
    return px / 1.333


def _px_to_col_width(px: int) -> float:
    """像素转 Excel 列宽字符数（Calibri 11 默认字体的近似换算公式）。"""
    return (px - 5) / 7


def _to_plain_text(value) -> str:
    """写入单元格前的兜底拍平：防止上游意外返回 tuple/list/嵌套 dict 时把结构体写入 Excel。"""
    if value is None:
        return ""
    if isinstance(value, (tuple, list)):
        value = value[0] if value else ""
    if isinstance(value, dict):
        value = value.get("translated") or value.get("text") or value.get("value") or value.get("original") or ""
    return str(value).strip()


def write_to_template(data_list, template_path, output_path):
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"找不到模板文件: {template_path}")

    wb = openpyxl.load_workbook(template_path)
    target_sheet = next((s for s in wb.sheetnames if s.startswith("总表")), None)
    if target_sheet is None:
        raise ValueError(f"模板中未找到以'总表'开头的工作表，当前 Sheet 列表: {wb.sheetnames}")

    ws = wb[target_sheet]
    ws.column_dimensions['E'].width = _px_to_col_width(BOX_W_PX)

    resized_tmp_files = []
    start_row = 5

    for data in data_list:
        ws.cell(row=start_row, column=2,  value=_to_plain_text(data.get("order_id", "")))
        ws.cell(row=start_row, column=6,  value=_to_plain_text(data.get("product_name", "")))
        ws.cell(row=start_row, column=10, value=_to_plain_text(data.get("fabric_quality", "")))
        ws.cell(row=start_row, column=11, value=_to_plain_text(data.get("color_print", "")))
        ws.cell(row=start_row, column=14, value=_to_plain_text(data.get("size_range", "")))

        image_path = data.get("image_path")
        if image_path and os.path.exists(image_path):
            try:
                resized_path, actual_w, actual_h = _resize_image(image_path)
                resized_tmp_files.append(resized_path)
                xl_img = XLImage(resized_path)
                xl_img.width  = actual_w
                xl_img.height = actual_h

                # 清空目标单元格原有文本，避免图片被拖走后底部残留文字
                ws.cell(row=start_row, column=5, value="")

                # 行高取 box 高度与图片实际高度的较大值，保证图片不被截断
                row_h_px = max(BOX_H_PX, actual_h)
                ws.row_dimensions[start_row].height = _px_to_row_height(row_h_px) + 4

                # 居中对齐：用单元格内剩余空间的一半作为偏移量
                col_off_px = max(0, (BOX_W_PX - actual_w) // 2)
                row_off_px = max(0, (row_h_px - actual_h) // 2)
                xl_img.anchor = OneCellAnchor(
                    _from=AnchorMarker(
                        col=4, colOff=pixels_to_EMU(col_off_px),
                        row=start_row - 1, rowOff=pixels_to_EMU(row_off_px),
                    ),
                    ext=XDRPositiveSize2D(pixels_to_EMU(actual_w), pixels_to_EMU(actual_h)),
                )
                ws.add_image(xl_img)
            except Exception:
                ws.row_dimensions[start_row].height = _px_to_row_height(BOX_H_PX)
        else:
            ws.row_dimensions[start_row].height = _px_to_row_height(BOX_H_PX)

        start_row += 1

    wb.save(output_path)

    for tmp in resized_tmp_files:
        try:
            os.remove(tmp)
        except OSError:
            pass
