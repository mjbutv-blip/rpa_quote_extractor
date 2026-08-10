import os
import re
import tempfile

import streamlit as st

from extractor_engine import (
    call_openai_to_extract,
    extract_fabric_quality_from_pdf,
    extract_size_range_from_pdf,
    extract_text_from_excel,
    extract_text_from_pdf,
    normalize_product_name_for_quote,
    postprocess_fabric_quality_for_quote,
)
from excel_writer import write_to_template

st.set_page_config(page_title="RPA 工艺单报价提取器", layout="wide")

st.title("服装工艺单 PDF / Excel 自动化提取与 RPA 填表系统")
st.write("上传多款产品的工艺单（PDF 或 Excel），AI 将自动提取核心字段并填入报价单总表模板中。")

api_key = st.sidebar.text_input(
    "输入 OpenAI API Key",
    type="password",
    value=os.environ.get("OPENAI_API_KEY", ""),
)

uploaded_files = st.file_uploader(
    "选择工艺单文件 (可多选，支持 PDF / Excel)",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True,
)
sort_mode = st.selectbox(
    "输出排序方式",
    ["按上传顺序", "按文件名/订单号自动排序", "手动指定顺序"],
)
custom_order_text = ""
if sort_mode == "手动指定顺序":
    custom_order_text = st.text_area(
        "手动顺序（按顺序输入订单号或文件名片段，每行一个，也可用逗号分隔）",
        placeholder="2114315\n2114322\n2114330",
        height=100,
    )
template_path = "templates/报价单总表模板.xlsx"


def _natural_sort_key(filename: str):
    parts = re.split(r"(\d+)", filename.lower())
    return [int(part) if part.isdigit() else part for part in parts]


def _order_id_from_filename(filename: str) -> str:
    match = re.search(r"\d{7}", filename)
    return match.group(0) if match else ""


def _parse_custom_order(value: str):
    return [item.strip().lower() for item in re.split(r"[\n,，;；]+", value or "") if item.strip()]


def _custom_order_rank(custom_order, order_id: str = "", filename: str = ""):
    haystacks = [str(order_id or "").lower(), str(filename or "").lower()]
    for idx, token in enumerate(custom_order):
        if any(token == haystack or token in haystack for haystack in haystacks):
            return idx
    return len(custom_order)


def _file_sort_key(file, sort_mode: str, custom_order):
    if sort_mode == "按文件名/订单号自动排序":
        return _natural_sort_key(file.name)
    if sort_mode == "手动指定顺序":
        return (
            _custom_order_rank(custom_order, _order_id_from_filename(file.name), file.name),
            _natural_sort_key(file.name),
        )
    return 0


def _result_sort_key(data, sort_mode: str, custom_order):
    filename = data.get("_source_filename", "")
    if sort_mode == "按文件名/订单号自动排序":
        return _natural_sort_key(filename)
    if sort_mode == "手动指定顺序":
        return (
            _custom_order_rank(custom_order, data.get("order_id", ""), filename),
            _natural_sort_key(filename),
        )
    return 0


def _ascii_safe(value) -> str:
    return str(value).encode("ascii", "backslashreplace").decode("ascii")


def _extract_color_print_from_text(text: str, fabric_quality: str = "") -> str:
    text_l = (text or "").lower()
    fabric = fabric_quality or ""
    is_multicoloured = "multicoloured" in text_l or "multi coloured" in text_l
    has_lace = "花边" in fabric
    has_logo_print = "logo" in text_l and ("rubber print" in text_l or "橡胶" in text)

    if is_multicoloured and has_lace:
        lines = ["大身前后片:双色花边", "其他材料:黑色配色"]
    elif is_multicoloured:
        lines = ["大身前后片:净色", "其他材料:配色"]
    else:
        lines = ["大身前后片:净色"]

    if has_logo_print:
        lines.append("身穿左下有单印印字logo")

    lines.extend([
        "",
        "1.定位问题:——",
        "2.印花朝向问题:——",
        "3.条纹/格子款式是否对条/对格:——",
    ])
    return "\n".join(lines)


if st.button("开始批量提取并生成报价单"):
    if not api_key:
        st.error("Please enter your OpenAI API Key.")
    elif not uploaded_files:
        st.warning("Please upload at least one file.")
    elif not os.path.exists(template_path):
        st.error(f"Template file not found: {template_path}")
    else:
        extracted_results = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        custom_order = _parse_custom_order(custom_order_text)
        files_to_process = sorted(
            list(uploaded_files),
            key=lambda file: _file_sort_key(file, sort_mode, custom_order),
        )

        for idx, uploaded_file in enumerate(files_to_process):
            status_text.text(f"Processing ({idx+1}/{len(files_to_process)}): {uploaded_file.name}...")
            file_ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
            is_pdf = file_ext == "pdf"

            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_ext}") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_file_path = tmp.name

            try:
                if is_pdf:
                    text = extract_text_from_pdf(tmp_file_path)
                else:
                    text = extract_text_from_excel(tmp_file_path)

                data = call_openai_to_extract(text, api_key)
                if is_pdf:
                    order_id = _order_id_from_filename(uploaded_file.name)
                    if order_id:
                        data["order_id"] = order_id
                    product_name = normalize_product_name_for_quote(text, data.get("product_name", ""))
                    if product_name:
                        data["product_name"] = product_name
                    try:
                        fabric_quality = extract_fabric_quality_from_pdf(tmp_file_path)
                        if fabric_quality:
                            data["fabric_quality"] = postprocess_fabric_quality_for_quote(
                                fabric_quality,
                                data.get("order_id", ""),
                            )
                    except Exception:
                        pass
                    color_print = _extract_color_print_from_text(text, data.get("fabric_quality", ""))
                    if color_print:
                        data["color_print"] = color_print
                    try:
                        size_range = extract_size_range_from_pdf(tmp_file_path)
                        if size_range:
                            data["size_range"] = size_range
                    except Exception:
                        pass
                data["_source_filename"] = uploaded_file.name
                extracted_results.append(data)
                file_type_label = "PDF" if is_pdf else "Excel"
                st.success(f"[{file_type_label}] {uploaded_file.name} extraction succeeded.")

            except Exception as e:
                st.error(f"{uploaded_file.name} extraction failed: {_ascii_safe(e)}")
            finally:
                if os.path.exists(tmp_file_path):
                    os.remove(tmp_file_path)

            progress_bar.progress((idx + 1) / len(files_to_process))

        status_text.text("All files processed. Writing Excel...")

        if extracted_results:
            extracted_results.sort(key=lambda data: _result_sort_key(data, sort_mode, custom_order))
            output_path = "templates/生成的报价总表_output.xlsx"
            try:
                write_to_template(extracted_results, template_path, output_path)
                st.success("Quote workbook generated successfully.")

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="Download generated quote workbook",
                        data=f,
                        file_name="生成的报价单总表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as e:
                st.error(f"Excel write failed: {_ascii_safe(e)}")
