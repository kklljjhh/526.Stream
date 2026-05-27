import streamlit as st
from processor import process_multiple_excel, process_excel

st.set_page_config(page_title="Excel 批量处理器", layout="centered")
st.title("📊 Excel 批量数据处理工具")
st.markdown("支持同时上传多个 `.xls` / `.xlsx` 文件，自动打包下载。")

uploaded_files = st.file_uploader(
    "选择文件（可多选）", 
    type=["xls", "xlsx"], 
    accept_multiple_files=True
)

if uploaded_files:
    files = [(f.getvalue(), f.name) for f in uploaded_files]
    
    with st.spinner(f"正在处理 {len(files)} 个文件..."):
        try:
            if len(files) == 1:
                # 单文件直接下载 xlsx
                result = process_excel(files[0][0], files[0][1])
                st.success("处理完成！")
                st.download_button(
                    label="📥 下载处理后的文件",
                    data=result,
                    file_name=files[0][1].rsplit(".", 1)[0] + "_处理后.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                # 多文件打包成 ZIP
                zip_bytes = process_multiple_excel(files)
                st.success(f"全部处理完成！共 {len(files)} 个文件")
                st.download_button(
                    label="📥 下载批量处理结果（ZIP）",
                    data=zip_bytes,
                    file_name="处理后数据.zip",
                    mime="application/zip"
                )
        except Exception as e:
            st.error(f"处理失败：{str(e)}")
