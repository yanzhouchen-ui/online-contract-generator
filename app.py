import streamlit as st
import pandas as pd
from docx import Document
import os
import zipfile
import io
from datetime import datetime

from utils.docx_utils import format_date_spanish, replace_text_in_doc

# 配置页面
st.set_page_config(page_title="线上合同制作器", page_icon="📄")

# 获取内置模板列表（使用绝对路径）
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")

def get_builtin_templates():
    """获取内置模板列表"""
    if not os.path.exists(TEMPLATES_DIR):
        return []
    templates = [f for f in os.listdir(TEMPLATES_DIR) 
                 if f.endswith('.docx') and not f.startswith('~$')]
    return sorted(templates)


def main():
    st.title("📄 线上合同制作器")
    st.write("上传 Excel 数据，批量生成合同文件。")
    
    # 1. 上传 Excel 文件
    st.subheader("1. 上传 Excel 数据文件")
    uploaded_excel = st.file_uploader(
        "选择 Excel 文件",
        type=['xlsx'],
        help="Excel 文件应包含合同数据，每行代表一个合同"
    )
    
    # 2. 选择模板
    st.subheader("2. 选择合同模板")
    
    builtin_templates = get_builtin_templates()
    template_options = ["使用所有内置模板", "上传自定义模板"]
    
    template_choice = st.radio(
        "模板来源",
        template_options,
        horizontal=True
    )
    
    uploaded_templates = None
    
    if template_choice == "使用所有内置模板":
        if builtin_templates:
            st.info(f"将使用以下 {len(builtin_templates)} 个内置模板生成合同：")
            for tmpl in builtin_templates:
                st.text(f"  • {tmpl}")
        else:
            st.error("未找到内置模板，请上传自定义模板")
    else:
        uploaded_templates = st.file_uploader(
            "上传 Word 模板（可多选）",
            type=['docx'],
            accept_multiple_files=True,
            help="上传一个或多个 .docx 格式的合同模板"
        )
        if uploaded_templates:
            st.info(f"将使用上传的 {len(uploaded_templates)} 个模板生成合同")
    
    # 3. 配置选项
    st.subheader("3. 配置选项")
    
    col1, col2 = st.columns(2)
    with col1:
        selected_date = st.date_input(
            "签署日期",
            datetime.now(),
            help="合同签署日期，将格式化为西班牙语格式"
        )
    
    with col2:
        price_val = st.text_input(
            "合同价格",
            value="0.25",
            help="替换模板中的 [precio] 字段"
        )
    
    # 显示日期预览
    spanish_date_str = format_date_spanish(selected_date)
    st.info(f"生成的日期格式: **{spanish_date_str}**")
    
    # 4. 生成合同
    st.subheader("4. 生成合同")
    
    # 检查是否可以生成
    can_generate = uploaded_excel is not None
    if template_choice == "使用所有内置模板":
        can_generate = can_generate and len(builtin_templates) > 0
    else:
        can_generate = can_generate and uploaded_templates is not None and len(uploaded_templates) > 0
    
    if not can_generate:
        st.warning("请上传 Excel 文件并选择模板")
        return
    
    if st.button("开始生成合同", type="primary", disabled=not can_generate):
        try:
            # 读取 Excel
            df = pd.read_excel(uploaded_excel)
            st.success(f"成功读取 Excel，共 {len(df)} 条数据")
            
            # 显示数据预览
            with st.expander("查看数据预览"):
                st.dataframe(df.head())
            
            # 准备生成
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # 确定要使用的模板列表
            if template_choice == "使用所有内置模板":
                templates_to_use = [(name, os.path.join(TEMPLATES_DIR, name)) for name in builtin_templates]
            else:
                # 使用上传的模板
                templates_to_use = [(tmpl.name, tmpl) for tmpl in uploaded_templates]
            
            total_templates = len(templates_to_use)
            total_items = len(df) * total_templates
            processed = 0
            
            # 创建内存中的 ZIP 文件
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                
                for idx, row in df.iterrows():
                    # 准备替换字典
                    replacements = {}
                    
                    # 添加 Excel 数据: [Column Name] -> Value
                    for col in df.columns:
                        val = row[col]
                        # 处理空值和非字符串
                        if pd.isna(val):
                            val = ""
                        else:
                            val = str(val)
                        # 去除列名空格，防止匹配失败
                        clean_col = " ".join(str(col).split())
                        
                        # 添加多种大小写形式，确保匹配成功
                        replacements[f"[{clean_col}]"] = val
                        replacements[f"[{clean_col.lower()}]"] = val
                        replacements[f"[{clean_col.upper()}]"] = val
                    
                    # 添加日期字段
                    replacements["[Fecha]"] = spanish_date_str
                    replacements["[fecha]"] = spanish_date_str
                    replacements["[Date]"] = spanish_date_str
                    
                    # 添加价格字段
                    replacements["[precio]"] = price_val
                    replacements["[Precio]"] = price_val
                    
                    # 确定文件夹名称 (使用第一列)
                    identifier = str(row[df.columns[0]]).strip()
                    identifier = "".join([c for c in identifier if c.isalnum() or c in (' ', '-', '_')]).strip()
                    
                    # 遍历所有模板
                    for template_name, template_source in templates_to_use:
                        status_text.text(f"正在处理: {identifier} - {template_name} ({idx + 1}/{len(df)})")
                        
                        # 加载模板
                        if isinstance(template_source, str):
                            # 内置模板（文件路径）
                            doc = Document(template_source)
                        else:
                            # 上传的模板（BytesIO）
                            template_source.seek(0)
                            doc = Document(template_source)
                        
                        # 替换文本
                        replace_text_in_doc(doc, replacements)
                        
                        # 保存到内存流
                        doc_io = io.BytesIO()
                        doc.save(doc_io)
                        doc_io.seek(0)
                        
                        # 添加到 ZIP
                        zip_file.writestr(f"{identifier}/{template_name}", doc_io.read())
                        
                        processed += 1
                        progress_bar.progress(processed / total_items)
            
            # 完成
            progress_bar.progress(100)
            status_text.text("生成完成！")
            st.balloons()
            
            # 提供下载按钮
            st.download_button(
                label="📥 下载所有合同 (ZIP)",
                data=zip_buffer.getvalue(),
                file_name=f"Contracts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )
            
        except Exception as e:
            st.error(f"生成失败: {str(e)}")
            st.exception(e)


if __name__ == "__main__":
    main()
