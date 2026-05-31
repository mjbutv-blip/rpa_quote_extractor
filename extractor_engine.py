import os
import json
import tempfile
import pdfplumber
import fitz  # PyMuPDF
from anthropic import Anthropic


def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


def extract_image_from_pdf(pdf_path):
    """提取 PDF 第一页中面积最大的图片，保存为临时 PNG，返回绝对路径。无图片则返回 None。"""
    try:
        doc = fitz.open(pdf_path)
        page = doc[0]
        image_list = page.get_images(full=True)

        if not image_list:
            doc.close()
            return None

        # full=True 时各字段: (xref, smask, width, height, bpc, colorspace, ...)
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


def call_claude_to_extract(text, api_key):
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
        "  \"size_range\": \"String. Extract the size range / 尺码范围, e.g., '32/34; 36/38; 40/42'.\"\n"
        "}\n"
        "Respond ONLY with the JSON object. Do not include markdown formatting like ```json or any conversational text."
    )

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        system=system_prompt,
        messages=[
            {"role": "user", "content": f"Extract data from this text:\n\n{text}"}
        ]
    )

    res_text = response.content[0].text.strip()
    if res_text.startswith("```json"):
        res_text = res_text.split("```json")[1].split("```")[0].strip()
    elif res_text.startswith("```"):
        res_text = res_text.split("```")[1].split("```")[0].strip()

    return json.loads(res_text)
