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


def crop_garment_region(image_path: str, api_key: str) -> str:
    """用 Claude Vision 裁剪出服装款式草图区域，去掉下方做工说明表格。

    分两步：
    1. 让 Claude 找「做工表格从哪里开始」（只问一个数字，比四边更准确）
    2. 再让 Claude 给出款式图的左右边界（去掉两侧空白）
    失败时原样返回原图路径。
    """
    if not _PIL_AVAILABLE or not image_path or not os.path.exists(image_path):
        return image_path
    try:
        with open(image_path, "rb") as f:
            img_bytes = f.read()

        img = PILImage.open(io.BytesIO(img_bytes)).convert("RGB")
        img_w, img_h = img.size

        ext  = image_path.rsplit(".", 1)[-1].lower()
        mime = _ext_to_mime(ext) or "image/png"
        b64  = base64.b64encode(img_bytes).decode("utf-8")

        client = Anthropic(api_key=api_key)

        # ── 步骤1：找做工说明表格的起始位置（垂直分界线）──────────
        r1 = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=120,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": mime, "data": b64}},
                    {
                        "type": "text",
                        "text": (
                            "This is an apparel tech pack. It has a garment sketch in the UPPER part "
                            "and specification tables (FABRIC / SIZE / PRINT / OBSERVATIONS or similar labels) "
                            "in the LOWER part.\n\n"
                            "At what percentage from the TOP of the image do the specification tables begin? "
                            "(0% = very top, 100% = very bottom)\n\n"
                            "If there are no tables, answer 95.\n"
                            "Return ONLY a JSON: {\"table_start_pct\": <number>}"
                        ),
                    },
                ],
            }],
        )
        raw1 = r1.content[0].text.strip()
        if "```" in raw1:
            raw1 = raw1.split("```")[1].split("```")[0].strip()
            if raw1.startswith("json"):
                raw1 = raw1[4:].strip()
        table_start = float(json.loads(raw1)["table_start_pct"])
        # 保守处理：稍微上移 2%，确保不带进表格顶部的边框线
        sketch_bottom = max(5.0, min(95.0, table_start - 2.0))

        # ── 步骤2：找款式图左右边界（水平居中裁剪）──────────────────
        r2 = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=120,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": mime, "data": b64}},
                    {
                        "type": "text",
                        "text": (
                            f"Focus ONLY on the upper {sketch_bottom:.0f}% of this image (the garment sketch area). "
                            "What are the leftmost and rightmost extents of the garment drawing itself "
                            "(ignore blank white margins)?\n"
                            "Return ONLY JSON: {\"left_pct\": <0-100>, \"right_pct\": <0-100>}"
                        ),
                    },
                ],
            }],
        )
        raw2 = r2.content[0].text.strip()
        if "```" in raw2:
            raw2 = raw2.split("```")[1].split("```")[0].strip()
            if raw2.startswith("json"):
                raw2 = raw2[4:].strip()
        lr = json.loads(raw2)
        left_pct  = max(0.0,   float(lr["left_pct"])  - 2.0)
        right_pct = min(100.0, float(lr["right_pct"]) + 2.0)

        # ── 裁剪 ────────────────────────────────────────────────────
        x0 = int(left_pct        / 100 * img_w)
        y0 = 0
        x1 = int(right_pct       / 100 * img_w)
        y1 = int(sketch_bottom   / 100 * img_h)

        if (x1 - x0) < 20 or (y1 - y0) < 20:
            return image_path

        cropped = img.crop((x0, y0, x1, y1))
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".png")
        os.close(tmp_fd)
        cropped.save(tmp_path, format="PNG")
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
