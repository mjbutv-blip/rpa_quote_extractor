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
    """将 PDF 第一页整体渲染为高清 PNG，返回临时文件路径。

    直接渲染整页（而非提取内嵌光栅图），可完整捕获矢量线稿、文字和图片，
    不会因为线稿是矢量路径而丢失服装顶部/细节。
    """
    if not _FITZ_AVAILABLE:
        return None
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        # 2× 缩放保证清晰度
        mat = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        doc.close()
        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".png")
        os.close(tmp_fd)
        pix.save(tmp_path)
        return tmp_path
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


def _sketch_score(img_bytes: bytes) -> float:
    """给图片打分，分数越高越像服装线稿。

    核心逻辑：线稿 = 白色背景（>45%）+ 图片面积大。
    - 白色背景比例是主要过滤条件（排除时装照、面料纹理等深色图片）
    - 面积是主要排序依据（线稿图比 Logo、文字框大得多）
    - 完全空白的图（无任何深色内容）排除
    """
    try:
        img = PILImage.open(io.BytesIO(img_bytes)).convert("L")
        pixels = list(img.getdata())
        total = len(pixels)
        white = sum(1 for p in pixels if p > 210) / total
        dark  = sum(1 for p in pixels if p < 100) / total
        area  = img.width * img.height
        if white < 0.45 or dark < 0.005:   # 过滤深色图片和空白图
            return 0.0
        return white * area                  # 白色比例 × 面积
    except Exception:
        return 0.0


def extract_style_image_from_excel(excel_path: str):
    """从 Excel 嵌入图片中挑选最像服装线稿的图，保存为临时 PNG，返回路径或 None。

    选图标准：白色背景占比高 × 深色线条存在 × 面积大（加权得分最高者）。
    这样可区分线稿（白底线条图）与时装照（深色背景）或面料图（均匀纹理）。
    """
    if not _PIL_AVAILABLE:
        return None
    if not excel_path.lower().endswith((".xlsx", ".xlsm")):
        return None
    try:
        with zipfile.ZipFile(excel_path, "r") as z:
            media_files = sorted(
                f for f in z.namelist() if f.startswith("xl/media/") and "." in f.rsplit("/", 1)[-1]
            )
            best_score, best_bytes = 0.0, None
            for media_file in media_files:
                ext = media_file.rsplit(".", 1)[-1].lower()
                if _ext_to_mime(ext) is None:
                    continue
                img_bytes = z.read(media_file)
                score = _sketch_score(img_bytes)
                if score > best_score:
                    best_score = score
                    best_bytes = img_bytes
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
    """用 OpenCV 检测三区域边界，裁剪出中间的服装款式草图。

    工艺单图片通常分三段：
      [顶部表头] → [中间款式图] → [底部做工说明表格]
    本函数同时检测顶部和底部边界，只保留中间款式图区域。
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

        # 二值化：深色内容（文字/线条）→ 白色；背景 → 黑色
        _, binary = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY_INV)

        # ── 表头检测用「严格宽核」(≥75% 图宽)：只有真正全宽的表格线才触发 ──
        # 服装的领口/肩带/接缝线宽度远达不到 75%，不会被误判为表头
        header_line_w = max(10, w * 3 // 4)
        hdr_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (header_line_w, 1))
        hdr_lines  = cv2.morphologyEx(binary, cv2.MORPH_OPEN, hdr_kernel)

        # ── 表格检测用「宽松宽核」(≥25% 图宽)：表格内部格线也能被找到 ──
        table_line_w = max(10, w // 4)
        tbl_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (table_line_w, 1))
        tbl_lines  = cv2.morphologyEx(binary, cv2.MORPH_OPEN, tbl_kernel)

        # ── 步骤1：找顶部表头底边（sketch_top）────────────────────────
        # 表头线必须：① 跨越 75%+ 图宽  ② 位于顶部 35% 区域  ③ 至少 2 条（单条可能是服装线）
        header_zone = int(h * 0.35)
        hdr_rows = np.where(np.any(hdr_lines[:header_zone], axis=1))[0]
        if len(hdr_rows) >= 2:
            sketch_top_y = int(hdr_rows[-1]) + 4   # 多条宽线 → 确认是表头
        else:
            sketch_top_y = 0                         # 零或一条 → 可能是服装线，不裁顶部

        # ── 步骤2：找底部表格顶边（sketch_bottom）─────────────────────
        search_from = max(sketch_top_y + 20, int(h * 0.30))
        tbl_rows = np.where(np.any(tbl_lines, axis=1))[0]
        bottom_lines = tbl_rows[tbl_rows >= search_from]
        table_top_y = int(bottom_lines[0]) if len(bottom_lines) > 0 else h

        # 行密度分析兜底：没有明显边框线时，用文字密度判断底部表格起点
        if table_top_y == h:
            row_density = np.sum(binary, axis=1).astype(float) / (w * 255)
            win = max(3, h // 40)
            density_smooth = np.convolve(row_density, np.ones(win) / win, mode='same')
            TEXT_THRESH, CONSEC = 0.10, 4
            count = 0
            for y in range(search_from, h):
                count = count + 1 if density_smooth[y] > TEXT_THRESH else 0
                if count >= CONSEC:
                    table_top_y = y - CONSEC + 1
                    break

        # 若底部分界线在底部 10% 内，视为无表格
        if table_top_y > int(h * 0.90):
            table_top_y = h

        sketch_top  = sketch_top_y
        sketch_bottom = max(sketch_top + 20, table_top_y - 4)

        # ── 步骤3：列投影找左右边界（去空白边距）─────────────────────
        sketch_zone = binary[sketch_top:sketch_bottom, :]
        col_sums = np.sum(sketch_zone, axis=0)
        content_cols = np.where(col_sums > 0)[0]
        if len(content_cols) > 0:
            left  = max(0, int(content_cols[0])  - 8)
            right = min(w, int(content_cols[-1]) + 8)
        else:
            left, right = 0, w

        # ── 裁剪并保存 ───────────────────────────────────────────────
        cropped = img[sketch_top:sketch_bottom, left:right]
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
