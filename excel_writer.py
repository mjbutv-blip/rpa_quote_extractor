import os
import tempfile
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

# 款式图在单元格内的最大显示尺寸（像素）
MAX_IMG_PX = 160


def _resize_image(src_path: str) -> tuple:
    """等比例缩放图片至最大 MAX_IMG_PX，返回 (新临时文件路径, 实际宽px, 实际高px)。"""
    img = PILImage.open(src_path).convert("RGBA")
    img.thumbnail((MAX_IMG_PX, MAX_IMG_PX), PILImage.LANCZOS)
    actual_w, actual_h = img.size
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".png")
    os.close(tmp_fd)
    img.save(tmp_path, format="PNG", optimize=False)
    return tmp_path, actual_w, actual_h


def _px_to_row_height(px: int) -> float:
    """像素转 Excel 行高（点单位，1pt ≈ 1.333px @ 96dpi）。"""
    return px / 1.333


def write_to_template(data_list, template_path, output_path):
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"找不到模板文件: {template_path}")

    wb = openpyxl.load_workbook(template_path)
    target_sheet = next((s for s in wb.sheetnames if s.startswith("总表")), None)
    if target_sheet is None:
        raise ValueError(f"模板中未找到以'总表'开头的工作表，当前 Sheet 列表: {wb.sheetnames}")

    ws = wb[target_sheet]
    ws.column_dimensions['E'].width = 14

    resized_tmp_files = []
    start_row = 5

    for data in data_list:
        ws.cell(row=start_row, column=2,  value=data.get("order_id", ""))
        ws.cell(row=start_row, column=6,  value=data.get("product_name", ""))
        ws.cell(row=start_row, column=10, value=data.get("fabric_quality", ""))
        ws.cell(row=start_row, column=11, value=data.get("color_print", ""))
        ws.cell(row=start_row, column=14, value=data.get("size_range", ""))

        image_path = data.get("image_path")
        if image_path and os.path.exists(image_path):
            try:
                resized_path, actual_w, actual_h = _resize_image(image_path)
                resized_tmp_files.append(resized_path)
                xl_img = XLImage(resized_path)
                # 用实际缩放后的尺寸，保持原始宽高比，不拉伸
                xl_img.width  = actual_w
                xl_img.height = actual_h
                ws.add_image(xl_img, f"E{start_row}")
                # 行高按图片实际高度设置，保证图片不被截断
                ws.row_dimensions[start_row].height = _px_to_row_height(actual_h) + 4
            except Exception:
                ws.row_dimensions[start_row].height = _px_to_row_height(MAX_IMG_PX)
        else:
            ws.row_dimensions[start_row].height = _px_to_row_height(MAX_IMG_PX)

        start_row += 1

    wb.save(output_path)

    for tmp in resized_tmp_files:
        try:
            os.remove(tmp)
        except OSError:
            pass
