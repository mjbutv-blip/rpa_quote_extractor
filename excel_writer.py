import os
import tempfile
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage

MAX_IMG_PX = 80


def _resize_image(src_path: str) -> str:
    """等比例缩放图片至最大 80x80 px，返回新临时文件路径。"""
    img = PILImage.open(src_path).convert("RGBA")
    img.thumbnail((MAX_IMG_PX, MAX_IMG_PX), PILImage.LANCZOS)
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".png")
    os.close(tmp_fd)
    img.save(tmp_path, format="PNG")
    return tmp_path


def write_to_template(data_list, template_path, output_path):
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"找不到模板文件: {template_path}")

    wb = openpyxl.load_workbook(template_path)
    target_sheet = next((s for s in wb.sheetnames if s.startswith("总表")), None)
    if target_sheet is None:
        raise ValueError(f"模板中未找到以'总表'开头的工作表，当前 Sheet 列表: {wb.sheetnames}")

    ws = wb[target_sheet]

    # E 列列宽固定设置一次
    ws.column_dimensions['E'].width = 12

    resized_tmp_files = []
    start_row = 5  # 强制从第 5 行开始，直接覆盖占位符

    for data in data_list:
        # 覆盖写入文本字段
        ws.cell(row=start_row, column=2,  value=data.get("order_id", ""))
        ws.cell(row=start_row, column=6,  value=data.get("product_name", ""))
        ws.cell(row=start_row, column=10, value=data.get("fabric_quality", ""))
        ws.cell(row=start_row, column=11, value=data.get("color_print", ""))
        ws.cell(row=start_row, column=14, value=data.get("size_range", ""))

        # 插入款式图到 E 列
        image_path = data.get("image_path")
        if image_path and os.path.exists(image_path):
            try:
                resized_path = _resize_image(image_path)
                resized_tmp_files.append(resized_path)
                xl_img = XLImage(resized_path)
                xl_img.width = MAX_IMG_PX
                xl_img.height = MAX_IMG_PX
                ws.add_image(xl_img, f"E{start_row}")
            except Exception:
                pass  # 图片插入失败不阻断文字写入

        # 设置当前行高
        ws.row_dimensions[start_row].height = 80
        start_row += 1

    wb.save(output_path)

    for tmp in resized_tmp_files:
        try:
            os.remove(tmp)
        except OSError:
            pass
