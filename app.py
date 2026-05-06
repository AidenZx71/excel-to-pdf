import streamlit as st
import pandas as pd
from weasyprint import HTML
import io
import zipfile

st.set_page_config(page_title="Excel 转 PDF 神器", page_icon="📄")

st.title("📄 Excel 批量转 PDF 小工具")
st.markdown("上传一个包含多个 Sheet 的 Excel 文件，系统会自动将每个 Sheet 转换成单独的 PDF，并打包成 ZIP 供你下载。")

if 'zip_data' not in st.session_state:
    st.session_state.zip_data = None

uploaded_file = st.file_uploader("请在此上传 Excel 文件 (.xlsx 或 .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.success(f"成功读取文件: {uploaded_file.name}")
    
    if st.button("🚀 开始转换", type="primary"):
        with st.spinner("正在拼命转换中，请稍候..."):
            try:
                xl = pd.ExcelFile(uploaded_file)
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    for sheet_name in xl.sheet_names:
                        # 【核心修复1】header=None 强制不把第一行识别为表头，保留最原始的结构
                        df = xl.parse(sheet_name, header=None)
                        
                        # 【核心修复2】把所有的缺失值 (NaN) 替换成空字符串，保持画面整洁
                        df = df.fillna("")
                        
                        # 【核心修复3】导出 HTML 时关闭 index 和 header，避免生成数字序号 0,1,2...
                        html_table = df.to_html(index=False, header=False)
                        
                        # 【核心修复4】在 CSS 中加入刚安装的中文字体 'WenQuanYi Zen Hei'
                        html_content = f"""
                        <html>
                        <head>
                            <style>
                                @page {{ size: A4 landscape; margin: 1cm; }}
                                body {{ 
                                    font-family: 'WenQuanYi Zen Hei', 'Microsoft YaHei', sans-serif; 
                                    font-size: 11px;
                                }}
                                table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                                td {{ border: 1px solid #ddd; padding: 6px; text-align: left; }}
                                /* 让包含关键字的行加粗加底色，模拟表头效果 */
                                tr:nth-child(2) td {{ background-color: #f8f9fa; font-weight: bold; }} 
                                h2 {{ color: #333; text-align: center; }}
                            </style>
                        </head>
                        <body>
                            <h2>数据表: {sheet_name}</h2>
                            {html_table}
                        </body>
                        </html>
                        """
                        
                        pdf_bytes = HTML(string=html_content).write_pdf()
                        zip_file.writestr(f"{sheet_name}.pdf", pdf_bytes)
                
                st.session_state.zip_data = zip_buffer.getvalue()
                
            except Exception as e:
                st.error(f"转换过程中出现错误: {e}")

    if st.session_state.zip_data is not None:
        st.success("🎉 转换成功！请点击下方按钮下载。")
        st.download_button(
            label="⬇️ 下载生成的 ZIP 压缩包",
            data=st.session_state.zip_data,
            file_name="Converted_PDFs.zip",
            mime="application/zip"
        )
