import os
import re
import tempfile
import uuid

import streamlit as st
try:
    from streamlit_sortables import sort_items
except ImportError:
    sort_items = None

from extractor_engine import (
    call_openai_to_extract,
    extract_fabric_quality_from_pdf,
    extract_images_from_excel,
    extract_size_range_from_pdf,
    extract_text_from_excel,
    extract_text_from_pdf,
    normalize_product_name_for_quote,
    postprocess_fabric_quality_for_quote,
)
from customer_templates import available_template_options
from el_quote_rules import apply_el_quote_rules
from excel_writer import write_to_template

st.set_page_config(page_title="RPA 工艺单报价提取器", layout="wide")

st.title("服装工艺单 PDF / Excel 自动化提取与 RPA 填表系统")
st.write("上传多款产品的工艺单（PDF 或 Excel），AI 将自动提取核心字段并填入报价单总表模板中。")

api_key = st.sidebar.text_input(
    "输入 OpenAI API Key",
    type="password",
    value=os.environ.get("OPENAI_API_KEY", ""),
)

customer_id = st.selectbox(
    "选择客户",
    ["default", "NKD", "EL", "TK"],
    format_func=lambda value: {
        "default": "默认客户",
        "NKD": "NKD",
        "EL": "EL",
        "TK": "TK",
    }.get(value, value),
)
template_options = available_template_options(customer_id)
selected_template_index = st.selectbox(
    "选择报价单模板",
    list(range(len(template_options))),
    format_func=lambda idx: f"{template_options[idx]['customer_name']} / {template_options[idx]['template_path']}",
)
selected_customer_template = template_options[selected_template_index]
template_path = selected_customer_template["template_path"]
sample_quantity_default = selected_customer_template.get("sample_quantity_default", "")
quote_title = "" if customer_id == "default" else f"{customer_id}-报价单"
st.caption(
    f"报价单表头：{quote_title or '使用模板原表头'}；"
    f"样品数量默认值：{'NKD 默认话术' if sample_quantity_default else '留空'}"
)

uploaded_files = st.file_uploader(
    "选择工艺单文件 (可多选，支持 PDF / Excel)",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True,
)
sort_mode = st.selectbox(
    "输出排序方式",
    ["拖动排序", "按上传顺序", "按文件名/订单号自动排序", "手动指定顺序"],
)
dragged_order = []
if sort_mode == "拖动排序" and uploaded_files:
    if sort_items is None:
        st.warning("拖动排序组件未安装，当前会按上传顺序生成。")
    else:
        labels = [f"{idx + 1}. {file.name}" for idx, file in enumerate(uploaded_files)]
        dragged_order = sort_items(labels)

custom_order_text = ""
if sort_mode == "手动指定顺序":
    custom_order_text = st.text_area(
        "手动顺序（按顺序输入订单号或文件名片段，每行一个，也可用逗号分隔）",
        placeholder="2114315\n2114322\n2114330",
        height=100,
    )


def _new_task_id() -> str:
    return f"TASK-{uuid.uuid4().hex[:8].upper()}"


def _render_task_rows(container, rows):
    with container.container():
        st.subheader("后台任务")
        st.dataframe(rows, use_container_width=True, hide_index=True)


if st.session_state.get("quote_task_rows"):
    tasks_panel = st.empty()
    _render_task_rows(tasks_panel, st.session_state["quote_task_rows"])


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


def _drag_order_from_labels(labels):
    result = []
    for label in labels or []:
        filename = re.sub(r"^\d+\.\s*", "", label).strip()
        if filename:
            result.append(filename.lower())
    return result


def _file_sort_key(file, sort_mode: str, custom_order, drag_order):
    if sort_mode == "拖动排序":
        if not drag_order:
            return 0
        return (
            _custom_order_rank(drag_order, _order_id_from_filename(file.name), file.name),
            _natural_sort_key(file.name),
        )
    if sort_mode == "按文件名/订单号自动排序":
        return _natural_sort_key(file.name)
    if sort_mode == "手动指定顺序":
        return (
            _custom_order_rank(custom_order, _order_id_from_filename(file.name), file.name),
            _natural_sort_key(file.name),
        )
    return 0


def _result_sort_key(data, sort_mode: str, custom_order, drag_order):
    filename = data.get("_source_filename", "")
    if sort_mode == "拖动排序":
        if not drag_order:
            return 0
        return (
            _custom_order_rank(drag_order, data.get("order_id", ""), filename),
            _natural_sort_key(filename),
        )
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
    has_body_print = any(keyword in text_l for keyword in (
        "all over print",
        "aop",
        "print",
        "printed",
        "pattern",
        "floral",
        "flower",
        "stripe",
        "striped",
        "check",
        "checked",
        "plaid",
        "dot",
        "dotted",
        "animal print",
        "graphic",
    )) or any(keyword in text for keyword in (
        "印花",
        "花型",
        "花版",
        "满印",
        "图案",
        "条纹",
        "格子",
        "波点",
        "碎花",
        "花朵",
    ))

    if is_multicoloured and has_lace:
        lines = ["大身前后片:双色花边", "其他材料:黑色配色"]
    elif has_body_print and not has_logo_print:
        lines = ["大身前后片:印花"]
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
        drag_order = _drag_order_from_labels(dragged_order)
        files_to_process = sorted(
            list(uploaded_files),
            key=lambda file: _file_sort_key(file, sort_mode, custom_order, drag_order),
        )
        st.session_state["quote_task_rows"] = [
            {
                "任务ID": _new_task_id(),
                "文件名": file.name,
                "客户": customer_id,
                "模板": selected_customer_template["customer_name"],
                "状态": "等待中",
                "错误": "",
            }
            for file in files_to_process
        ]
        tasks_panel = st.empty()
        _render_task_rows(tasks_panel, st.session_state["quote_task_rows"])

        for idx, uploaded_file in enumerate(files_to_process):
            status_text.text(f"Processing ({idx+1}/{len(files_to_process)}): {uploaded_file.name}...")
            st.session_state["quote_task_rows"][idx]["状态"] = "处理中"
            _render_task_rows(tasks_panel, st.session_state["quote_task_rows"])
            file_ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
            is_pdf = file_ext == "pdf"

            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_ext}") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_file_path = tmp.name

            try:
                images = None
                if is_pdf:
                    text = extract_text_from_pdf(tmp_file_path)
                else:
                    text = extract_text_from_excel(tmp_file_path)
                    images = extract_images_from_excel(tmp_file_path)

                data = call_openai_to_extract(text, api_key, images=images)
                order_id = _order_id_from_filename(uploaded_file.name)
                if order_id:
                    data["order_id"] = order_id
                product_name = normalize_product_name_for_quote(text, data.get("product_name", ""))
                if product_name:
                    data["product_name"] = product_name
                color_print = _extract_color_print_from_text(text, data.get("fabric_quality", ""))
                if color_print:
                    data["color_print"] = color_print
                if is_pdf:
                    try:
                        fabric_quality = extract_fabric_quality_from_pdf(tmp_file_path)
                        if fabric_quality:
                            data["fabric_quality"] = postprocess_fabric_quality_for_quote(
                                fabric_quality,
                                data.get("order_id", ""),
                            )
                    except Exception:
                        pass
                    try:
                        size_range = extract_size_range_from_pdf(tmp_file_path)
                        if size_range:
                            data["size_range"] = size_range
                    except Exception:
                        pass
                if customer_id == "EL":
                    data = apply_el_quote_rules(text, data)
                data["_source_filename"] = uploaded_file.name
                extracted_results.append(data)
                file_type_label = "PDF" if is_pdf else "Excel"
                st.session_state["quote_task_rows"][idx]["状态"] = "报价提取完成"
                st.success(f"[{file_type_label}] {uploaded_file.name} extraction succeeded.")

            except Exception as e:
                st.session_state["quote_task_rows"][idx]["状态"] = "失败"
                st.session_state["quote_task_rows"][idx]["错误"] = _ascii_safe(e)
                st.error(f"{uploaded_file.name} extraction failed: {_ascii_safe(e)}")
            finally:
                if os.path.exists(tmp_file_path):
                    os.remove(tmp_file_path)

            _render_task_rows(tasks_panel, st.session_state["quote_task_rows"])
            progress_bar.progress((idx + 1) / len(files_to_process))

        status_text.text("All files processed. Writing Excel...")

        if extracted_results:
            extracted_results.sort(key=lambda data: _result_sort_key(data, sort_mode, custom_order, drag_order))
            output_path = "templates/生成的报价总表_output.xlsx"
            try:
                write_to_template(
                    extracted_results,
                    template_path,
                    output_path,
                    sample_quantity_default=sample_quantity_default,
                    quote_title=quote_title,
                )
                for row in st.session_state["quote_task_rows"]:
                    if row["状态"] == "报价提取完成":
                        row["状态"] = "报价单已生成"
                _render_task_rows(tasks_panel, st.session_state["quote_task_rows"])
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
