import base64
import json
import os
import tempfile

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

# MIME types supported by the Anthropic vision API
_SUPPORTED_EXTS = {"jpeg", "jpg", "png", "gif", "webp"}


def _ext_to_mime(ext: str):
    ext = ext.lower().lstrip(".")
    if ext == "jpg":
        return "image/jpeg"
    if ext in _SUPPORTED_EXTS:
        return f"image/{ext}"
    return None


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


def extract_image_from_pdf(pdf_path: str):
    """提取 PDF 第一页中面积最大的图片，保存为临时 PNG，返回绝对路径。无图片则返回 None。"""
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
        xref = largest[0]

        base_image = doc.extract_image(xref)
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

        best_area = 0
        best_xref = None
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


def call_claude_to_extract(
    text: str,
    api_key: str,
    image_base64: str = None,
    media_type: str = "image/jpeg",
) -> dict:
    """调用 Claude 从工艺单文本（+ 可选款式图）中提取结构化字段。"""
    client = Anthropic(api_key=api_key)

    system_prompt = (
        "You are an expert data extractor for apparel tech packs. "
        "Analyze the provided text and extract the following 5 fields into a valid JSON object. "
        "Translate fields to Chinese where requested.\n"
        "JSON Schema:\n"
        "{\n"
        "  \"order_id\": \"String. Extract the order number / 款号 / 订单号, e.g., '2118879'.\",\n"
        "  \"product_name\": \"String. Extract the product name / 品名, and translate it to Chinese if it is in English/German, e.g., '无钢圈文胸'.\",\n"
        "  \"fabric_quality\": \"String. Extract the fabric composition and weight / 面料品质, translate to Chinese, e.g., '80%锦纶, 20%弹性纤维, 170 GSM'.\",\n"
        "  \"color_print\": \"String. Extract the color or print name / 颜色, translate to Chinese, e.g., '白色'.\",\n"
        "  \"size_range\": \"String. Extract ALL sizes from the size grading table or size list. "
        "Follow these rules strictly:\\n"
        "  1. Search for keywords: 'Size', 'Sizes', 'Size Range', 'Grading', '尺码', '规格', '号型', 'S/M/L', 'XS~XXL', or any size grid/table.\\n"
        "  2. Bra/lingerie band+cup sizes: output as 'band_cupA/band_cupB' per row, e.g. '70A/75A/80A/85A/90A' or '32A/34B/36B/38C'.\\n"
        "  3. Alpha sizes: output as 'XS/S/M/L/XL/XXL'.\\n"
        "  4. Numeric sizes: output as '36/38/40/42/44'.\\n"
        "  5. If the document uses a 2-axis size grid (e.g., band vs cup), list every unique size combination, e.g. '70A/70B/75A/75B/80A/80B'.\\n"
        "  6. Do NOT pick up fabric weights, order quantities, measurements in cm/inch, or price numbers as sizes.\\n"
        "  7. Separate all sizes with '/'. Include every size present — do not omit any.\"\n"
        "}\n"
        "【关键视觉校验】：客户提供的英文品名或描述可能存在误差。"
        "你必须同时观察提供的款式图片（如有）。"
        "如果文本描述（如 hipster）与图片实际显示的款式（如 女士三角裤）产生冲突，"
        "请【绝对优先】以图片显示的实际物理款式为准来提取/推断正确的【品名】字段，"
        "并翻译成最准确的中文服装专业术语。\n"
        "如果图片中包含尺码表或尺码网格，请同样以图片为准提取尺码，图片优先于文本。\n"
        "Respond ONLY with the JSON object. "
        "Do not include markdown formatting like ```json or any conversational text."
    )

    user_content = []
    if image_base64:
        user_content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": media_type,
                "data": image_base64,
            },
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
