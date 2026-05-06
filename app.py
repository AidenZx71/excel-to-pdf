import streamlit as st
import pandas as pd
from weasyprint import HTML
import io
import zipfile

st.set_page_config(page_title="Excel 转 PDF 神器", page_icon="📄")

st.title("📄 Excel 批量转 PDF 小工具")
st.markdown("上传一个包含多个 Sheet 的 Excel 文件，系统会自动将每个 Sheet 转换成单独的 PDF，并打包成 ZIP 供你下载。")

# 1. 文件上传组件
uploaded_file = st.file_uploader("请在此上传 Excel 文件 (.xlsx 或 .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success(f"成功读取文件: {uploaded_file.name}")
    
    if st.button("🚀 开始转换", type="primary"):
        with st.spinner("正在拼命转换中，请稍候..."):
            try:
                # 读取 Excel
                xl = pd.ExcelFile(uploaded_file)
                
                # 创建一个内存中的 ZIP 文件，用于打包多个 PDF
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    
                    # 遍历每个 Sheet
                    for sheet_name in xl.sheet_names:
                        df = xl.parse(sheet_name)
                        
                        # 转换成带样式的 HTML
                        html_table = df.to_html(index=False)
                        html_content = f"""
                        <html>
                        <head>
                            <style>
                                @page {{ size: A4 landscape; margin: 1cm; }}
                                body {{ font-family: sans-serif; }}
                                table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                                th, td {{ border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 12px; }}
                                th {{ background-color: #f2f2f2; font-weight: bold; }}
                                h2 {{ color: #333; text-align: center; }}
                            </style>
                        </head>
                        <body>
                            <h2>数据表: {sheet_name}</h2>
                            {html_table}
                        </body>
                        </html>
                        """
                        
                        # 使用 WeasyPrint 将 HTML 转换为 PDF 字节流
                        pdf_bytes = HTML(string=html_content).write_pdf()
                        
                        # 将生成的 PDF 写入内存中的 ZIP 文件
                        zip_file.writestr(f"{sheet_name}.pdf", pdf_bytes)
                
                # 准备下载
                st.success("🎉 转换成功！请点击下方按钮下载。")
                
                # 2. 下载按钮
                st.download_button(
                    label="⬇️ 下载生成的 ZIP 压缩包",
                    data=zip_buffer.getvalue(),
                    file_name="Converted_PDFs.zip",
                    mime="application/zip"
                )
                
            except Exception as e:
                st.error(f"转换过程中出现错误: {e}")