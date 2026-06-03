import base64
import io
import json
import os
import tempfile
import zipfile

import openpyxl
from anthropic import Anthropic

try:
    import fitz  # PyMuPDF
    _FITZ_AVAILABLE = True
except ImportError:
    _FITZ_AVAILABLE = False

try:
    import pdfplumber
    _PDFPLUMBER_AVAILABLE = True
except ImportError:
    _PDFPLUMBER_AVAILABLE = False

try:
    from PIL import Image as PILImage
    _PIL_AVAILABLE = True
except ImportError:
    _PIL_AVAILABLE = False

# MIME types supported by the Anthropic vision API
_SUPPORTED_EXTS = {"jpeg", "jpg", "png", "gif", "webp"}


def _ext_to_mime(ext: str):
    ext = ext.lower().lstrip(".")
    if ext == "jpg":
        return "image/jpeg"
    if ext in _SUPPORTED_EXTS:
        return f"image/{ext}"
    return None


# ── PDF 相关 ──────────────────────────────────────────────────────────────────

def extract_text_from_pdf(pdf_path: str) -> str:
    if not _PDFPLUMBER_AVAILABLE:
        raise RuntimeError("pdfplumber 未安装，无法处理 PDF 文件。")
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


def extract_image_from_pdf(pdf_path: str):
    """提取 PDF 第一页面积最大的图片，保存为临时 PNG，返回路径或 None。"""
    if not _FITZ_AVAILABLE:
        return None
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        image_list = page.get_images(full=True)
        if not image_list:
            doc.close()
            return None
        largest = max(image_list, key=lambda img: img[2] * img[3])
        base_image = doc.extract_image(largest[0])
        image_bytes = base_image["image"]
        doc.close()
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".png")
        os.close(tmp_fd)
        with open(tmp_path, "wb") as f:
            f.write(image_bytes)
        return os.path.abspath(tmp_path)
    except Exception:
        return None


def extract_image_base64(pdf_path: str, max_pages: int = 3):
    """扫描前 max_pages 页，找出面积最大的图片，返回 (base64_str, media_type) 或 None。"""
    if not _FITZ_AVAILABLE:
        return None
    try:
        doc = fitz.open(pdf_path)
        scan_pages = min(max_pages, len(doc))
        best_area, best_xref = 0, None
        for pn in range(scan_pages):
            for img in doc[pn].get_images(full=True):
                area = img[2] * img[3]
                if area > best_area:
                    best_area = area
                    best_xref = img[0]
        if best_xref is None:
            doc.close()
            return None
        base_image = doc.extract_image(best_xref)
        doc.close()
        mime = _ext_to_mime(base_image.get("ext", ""))
        if mime is None:
            return None
        b64 = base64.b64encode(base_image["image"]).decode("utf-8")
        return b64, mime
    except Exception:
        return None


# ── Excel 相关 ────────────────────────────────────────────────────────────────

def extract_text_from_excel(excel_path: str) -> str:
    """将 Excel 工艺单所有 sheet 的单元格内容转为纯文本，供 Claude 提取字段。"""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    lines = []
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        lines.append(f"=== Sheet: {sheet_name} ===")
        for row in ws.iter_rows(values_only=True):
            row_vals = [str(c) if c is not None else "" for c in row]
            row_text = "\t".join(row_vals)
            if row_text.strip():
                lines.append(row_text)
    return "\n".join(lines)


def extract_images_from_excel(excel_path: str) -> list:
    """从 .xlsx 文件中提取所有嵌入图片，返回 [(base64_str, media_type), ...] 列表。
    图片顺序：按文件名排序（image1, image2 ...）。
    """
    images = []
    if not excel_path.lower().endswith((".xlsx", ".xlsm")):
        return images
    try:
        with zipfile.ZipFile(excel_path, "r") as z:
            media_files = sorted(
                f for f in z.namelist() if f.startswith("xl/media/") and "." in f.rsplit("/", 1)[-1]
            )
            for media_file in media_files:
                ext = media_file.rsplit(".", 1)[-1].lower()
                mime = _ext_to_mime(ext)
                if mime is None:
                    continue
                img_bytes = z.read(media_file)
                b64 = base64.b64encode(img_bytes).decode("utf-8")
                images.append((b64, mime))
    except Exception:
        pass
    return images


def extract_style_image_from_excel(excel_path: str):
    """从 Excel 嵌入图片中挑选面积最大的作为款式图，保存为临时 PNG，返回路径或 None。"""
    if not _PIL_AVAILABLE:
        return None
    if not excel_path.lower().endswith((".xlsx", ".xlsm")):
        return None
    try:
        with zipfile.ZipFile(excel_path, "r") as z:
            media_files = sorted(
                f for f in z.namelist() if f.startswith("xl/media/") and "." in f.rsplit("/", 1)[-1]
            )
            best_area, best_bytes = 0, None
            for media_file in media_files:
                ext = media_file.rsplit(".", 1)[-1].lower()
                if _ext_to_mime(ext) is None:
                    continue
                img_bytes = z.read(media_file)
                try:
                    img = PILImage.open(io.BytesIO(img_bytes))
                    area = img.width * img.height
                    if area > best_area:
                        best_area = area
                        best_bytes = img_bytes
                except Exception:
                    pass
        if best_bytes:
            tmp_fd, tmp_path = tempfile.mkstemp(suffix=".png")
            os.close(tmp_fd)
            with open(tmp_path, "wb") as f:
                f.write(best_bytes)
            return tmp_path
    except Exception:
        pass
    return None


def crop_garment_region(image_path: str, api_key: str = None) -> str:
    """用 OpenCV 检测做工说明表格边界，裁剪出服装款式草图区域。

    完全基于图像处理，不调用 Claude API：
    1. 形态学水平线检测 → 找表格上边界
    2. 行密度分析兜底 → 找高密度文字区域起点
    3. 列投影 → 找款式图左右边界，去掉空白边距
    失败时原样返回原图路径。
    """
    if not image_path or not os.path.exists(image_path):
        return image_path
    try:
        import cv2
        import numpy as np

        img = cv2.imread(image_path, cv2.IMREAD_COLOR)
        if img is None:
            return image_path

        h, w = img.shape[:2]
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

        # ── 步骤1：形态学水平线检测（找表格边框线）─────────────────
        # 二值化：深色像素（文字/线条）变为白色
        _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

        # 水平线核：宽度至少占图片宽度的 1/4
        min_line_w = max(10, w // 4)
        h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (min_line_w, 1))
        h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel)

        # 在跳过顶部 30% 后，寻找第一条横跨全宽的线
        skip_y = int(h * 0.30)
        table_top_y = h  # 默认：无表格，不裁底部
        line_rows = np.where(np.any(h_lines[skip_y:], axis=1))[0]
        if len(line_rows) > 0:
            table_top_y = int(line_rows[0]) + skip_y

        # ── 步骤2：行密度分析兜底（处理无明显边框线的情况）─────────
        if table_top_y == h:
            row_density = np.sum(binary, axis=1).astype(float) / (w * 255)

            # 平滑后寻找高密度文字区（连续多行密度 > 阈值）
            win = max(3, h // 40)
            kernel_1d = np.ones(win) / win
            density_smooth = np.convolve(row_density, kernel_1d, mode='same')

            TEXT_THRESH = 0.10   # 行内超过 10% 像素为深色即判为文字行
            CONSEC_ROWS = 4      # 连续 N 行均满足才确认为表格区域

            count = 0
            for y in range(skip_y, h):
                if density_smooth[y] > TEXT_THRESH:
                    count += 1
                    if count >= CONSEC_ROWS:
                        table_top_y = y - CONSEC_ROWS + 1
                        break
                else:
                    count = 0

        # 如果分界线在图片底部 10% 以下，视为无表格，保留完整图片
        if table_top_y > int(h * 0.90):
            table_top_y = h

        # 留 4px 安全边距，避免把表格首行边框也带进去
        sketch_h = max(20, table_top_y - 4)

        # ── 步骤3：列投影找款式图左右边界（去空白边距）─────────────
        sketch_zone = binary[:sketch_h, :]
        col_sums = np.sum(sketch_zone, axis=0)
        content_cols = np.where(col_sums > 0)[0]

        if len(content_cols) > 0:
            left  = max(0, int(content_cols[0])  - 8)
            right = min(w, int(content_cols[-1]) + 8)
        else:
            left, right = 0, w

        # ── 裁剪并保存 ───────────────────────────────────────────────
        cropped = img[:sketch_h, left:right]
        if cropped.shape[0] < 20 or cropped.shape[1] < 20:
            return image_path

        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".png")
        os.close(tmp_fd)
        cv2.imwrite(tmp_path, cropped)
        return tmp_path

    except Exception:
        return image_path


# ── Claude 提取 ───────────────────────────────────────────────────────────────

def call_claude_to_extract(
    text: str,
    api_key: str,
    images: list = None,
) -> dict:
    """调用 Claude 从工艺单文本（+ 可选多张图片）中提取结构化字段。

    images: [(base64_str, media_type), ...] 或 None。
    可传入多张图片（款式图、面料色卡等），Claude 会综合所有图片提取信息。
    """
    client = Anthropic(api_key=api_key)

    system_prompt = (
        "You are an expert data extractor for apparel tech packs. "
        "You may receive multiple images: style/fashion photos AND fabric swatch or color sample images. "
        "Analyze ALL provided images together with the text to extract the following 5 fields into a valid JSON object.\n"
        "JSON Schema:\n"
        "{\n"
        "  \"order_id\": \"String. Extract the order number / 款号 / 订单号, e.g., '2118879'.\",\n"
        "  \"product_name\": \"String. Extract the product name / 品名, translate to Chinese, e.g., '无钢圈文胸'. "
        "If text and style image conflict, trust the image.\",\n"
        "  \"fabric_quality\": \"String. Extract the fabric composition and weight / 面料品质, translate to Chinese, "
        "e.g., '80%锦纶 20%弹性纤维 170GSM'. "
        "If fabric specs appear in an image (printed label, swatch tag, or fabric table in the image), extract from there too.\",\n"
        "  \"color_print\": \"String. Extract the color or print name / 颜色/印花, translate to Chinese, e.g., '白色' or '碎花印花'. "
        "If a color swatch or print image is provided, describe the color/print based on the image — this is the primary source.\",\n"
        "  \"size_range\": \"String. Extract ALL sizes from the size grading table or size list. "
        "Rules:\\n"
        "  1. Search for keywords: 'Size', 'Sizes', 'Size Range', 'Grading', '尺码', '规格', '号型', or any size grid/table.\\n"
        "  2. Bra/lingerie band+cup sizes: e.g. '70A/75A/80A/85A/90A' or '32A/34B/36B/38C'.\\n"
        "  3. Alpha sizes: e.g. 'XS/S/M/L/XL/XXL'.\\n"
        "  4. Numeric sizes: e.g. '36/38/40/42/44'.\\n"
        "  5. If a 2-axis size grid exists (band vs cup), list every combination, e.g. '70A/70B/75A/75B/80A/80B'.\\n"
        "  6. Do NOT pick up fabric weights, order quantities, cm/inch measurements, or prices as sizes.\\n"
        "  7. Separate all sizes with '/'. Include every size — do not omit any.\"\n"
        "}\n"
        "IMPORTANT: When multiple images are provided, treat them collectively:\n"
        "- Use style/garment images to verify product_name.\n"
        "- Use fabric swatch images or color sample images to determine fabric_quality and color_print.\n"
        "- If an image shows text (labels, tags, printed specs), read that text and use it for extraction.\n"
        "Respond ONLY with the JSON object. "
        "Do not include markdown formatting like ```json or any conversational text."
    )

    user_content = []
    if images:
        for b64, mime in images:
            user_content.append({
                "type": "image",
                "source": {"type": "base64", "media_type": mime, "data": b64},
            })
    user_content.append({
        "type": "text",
        "text": f"Extract data from this tech pack. Text content:\n\n{text}",
    })

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        system=system_prompt,
        messages=[{"role": "user", "content": user_content}],
    )

    res_text = response.content[0].text.strip()
    if res_text.startswith("```json"):
        res_text = res_text.split("```json")[1].split("```")[0].strip()
    elif res_text.startswith("```"):
        res_text = res_text.split("```")[1].split("```")[0].strip()

    return json.loads(res_text)
