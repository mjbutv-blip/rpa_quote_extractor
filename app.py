import os
import tempfile

import streamlit as st

from extractor_engine import (
    call_claude_to_extract,
    crop_garment_region,
    extract_image_base64,
    extract_image_from_pdf,
    extract_images_from_excel,
    extract_style_image_from_excel,
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

if st.button("🚀 开始批量提取并生成报价单"):
    if not api_key:
        st.error("请先在左侧输入您的 Anthropic API Key！")
    elif not uploaded_files:
        st.warning("请至少上传一个文件！")
    elif not os.path.exists(template_path):
        st.error(f"未找到模板文件：{template_path}，请将「报价单总表模板.xlsx」放入 templates/ 目录。")
    else:
        extracted_results = []
        image_tmp_files = []
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"正在处理 ({idx+1}/{len(uploaded_files)}): {uploaded_file.name}...")
            file_ext = uploaded_file.name.rsplit(".", 1)[-1].lower()
            is_pdf = file_ext == "pdf"

            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{file_ext}") as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_file_path = tmp.name

            try:
                if is_pdf:
                    # PDF：提取文本 + 单张最大图（多模态）
                    text = extract_text_from_pdf(tmp_file_path)
                    img_result = extract_image_base64(tmp_file_path)
                    images = [img_result] if img_result else None
                    data = call_claude_to_extract(text, api_key, images=images)
                    image_path = extract_image_from_pdf(tmp_file_path)
                else:
                    # Excel：提取文本 + 所有嵌入图片（多模态）
                    text = extract_text_from_excel(tmp_file_path)
                    excel_images = extract_images_from_excel(tmp_file_path)
                    images = excel_images if excel_images else None
                    data = call_claude_to_extract(text, api_key, images=images)
                    # 取面积最大的图作为输出模板中的款式图
                    image_path = extract_style_image_from_excel(tmp_file_path)

                # PDF 和 Excel 都可能提取到包含多种信息的合图，统一裁剪出款式草图区域
                if image_path:
                    cropped_path = crop_garment_region(image_path, api_key)
                    if cropped_path != image_path:
                        try:
                            os.remove(image_path)
                        except OSError:
                            pass
                        image_path = cropped_path

                data["image_path"] = image_path
                if image_path:
                    image_tmp_files.append(image_path)

                extracted_results.append(data)
                file_type_label = "PDF" if is_pdf else "Excel"
                img_count = len(images) if images else 0
                img_note = f"含 {img_count} 张图片 🖼️" if img_count else "无图片"
                st.success(f"✅ [{file_type_label}] {uploaded_file.name} 提取成功！（{img_note}）")

            except Exception as e:
                st.error(f"❌ {uploaded_file.name} 提取失败: {str(e)}")
            finally:
                if os.path.exists(tmp_file_path):
                    os.remove(tmp_file_path)

            progress_bar.progress((idx + 1) / len(uploaded_files))

        status_text.text("所有文件提取完毕！正在写入 Excel...")

        if extracted_results:
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
            finally:
                for img_path in image_tmp_files:
                    try:
                        if os.path.exists(img_path):
                            os.remove(img_path)
                    except OSError:
                        pass
