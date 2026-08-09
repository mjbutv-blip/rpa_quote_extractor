import os
import re
import tempfile

import streamlit as st

from extractor_engine import (
    call_claude_to_extract,
    extract_fabric_quality_from_pdf,
    extract_size_range_from_pdf,
    extract_text_from_excel,
    extract_text_from_pdf,
)
from excel_writer import write_to_template

st.set_page_config(page_title="RPA 工艺单报价提取器", layout="wide")

st.title("📂 服装工艺单 PDF / Excel 自动化提取与 RPA 填表系统")
st.write("上传多款产品的工艺单（PDF 或 Excel），AI 将自动提取核心字段并填入报价单总表模板中。")

api_key = st.sidebar.text_input(
    "输入 Anthropic API Key",
    type="password",
    value=os.environ.get("ANTHROPIC_API_KEY", ""),
)

uploaded_files = st.file_uploader(
    "选择工艺单文件 (可多选，支持 PDF / Excel)",
    type=["pdf", "xlsx", "xls"],
    accept_multiple_files=True,
)
template_path = "templates/报价单总表模板.xlsx"


def _natural_sort_key(filename: str):
    """按文件名中的数字自然排序，保证批量文件稳定写入顺序。"""
    parts = re.split(r"(\d+)", filename.lower())
    return [int(part) if part.isdigit() else part for part in parts]


def _order_id_from_filename(filename: str) -> str:
    match = re.search(r"\d{7}", filename)
    return match.group(0) if match else ""


if st.button("🚀 开始批量提取并生成报价单"):
    if not api_key:
        st.error("请先在左侧输入您的 Anthropic API Key！")
    elif not uploaded_files:
        st.warning("请至少上传一个文件！")
    elif not os.path.exists(template_path):
        st.error(f"未找到模板文件：{template_path}，请将「报价单总表模板.xlsx」放入 templates/ 目录。")
    else:
        extracted_results = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        files_to_process = sorted(uploaded_files, key=lambda f: _natural_sort_key(f.name))

        for idx, uploaded_file in enumerate(files_to_process):
            status_text.text(f"正在处理 ({idx+1}/{len(files_to_process)}): {uploaded_file.name}...")
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

                data = call_claude_to_extract(text, api_key)
                if is_pdf:
                    order_id = _order_id_from_filename(uploaded_file.name)
                    if order_id:
                        data["order_id"] = order_id
                    fabric_quality = extract_fabric_quality_from_pdf(tmp_file_path)
                    if fabric_quality:
                        data["fabric_quality"] = fabric_quality
                    size_range = extract_size_range_from_pdf(tmp_file_path)
                    if size_range:
                        data["size_range"] = size_range
                data["_source_filename"] = uploaded_file.name
                extracted_results.append(data)
                file_type_label = "PDF" if is_pdf else "Excel"
                st.success(f"✅ [{file_type_label}] {uploaded_file.name} 提取成功！")

            except Exception as e:
                st.error(f"❌ {uploaded_file.name} 提取失败: {str(e)}")
            finally:
                if os.path.exists(tmp_file_path):
                    os.remove(tmp_file_path)

            progress_bar.progress((idx + 1) / len(files_to_process))

        status_text.text("所有文件提取完毕！正在写入 Excel...")

        if extracted_results:
            extracted_results.sort(key=lambda data: _natural_sort_key(data.get("_source_filename", "")))
            output_path = "templates/生成的报价总表_output.xlsx"
            try:
                write_to_template(extracted_results, template_path, output_path)
                st.balloons()
                st.success("🎉 报价总表生成成功！")

                with open(output_path, "rb") as f:
                    st.download_button(
                        label="📥 点击下载最新生成的报价单总表.xlsx",
                        data=f,
                        file_name="生成的报价单总表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except Exception as e:
                st.error(f"写入 Excel 时发生错误: {str(e)}")
