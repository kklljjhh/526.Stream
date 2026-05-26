import streamlit as st
import zipfile
from io import BytesIO
from processor import process_excel

st.set_page_config(page_title="Excel 批量处理工具", layout="centered")

st.title("📊 Excel 批量处理工具")
st.markdown("支持一次上传多个 .xls / .xlsx 文件，自动批量处理后打包下载")

uploaded_files = st.file_uploader(
    "上传 Excel 文件（可多选）", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"已接收 {len(uploaded_files)} 个文件")
    
    if st.button("开始批量处理", type="primary"):
        with st.spinner("正在处理..."):
            results = []
            errors = []
            
            for uploaded in uploaded_files:
                try:
                    result_bytes = process_excel(uploaded.getvalue(), uploaded.name)
                    # 去掉原后缀，加 _处理后.xlsx
                    out_name = uploaded.name.rsplit('.', 1)[0] + "_处理后.xlsx"
                    results.append((out_name, result_bytes))
                except ValueError as e:
                    errors.append(f"{uploaded.name}: {e}")
                except Exception as e:
                    errors.append(f"{uploaded.name}: 处理出错 - {e}")
            
            # 显示错误
            if errors:
                for err in errors:
                    st.warning(err)
            
            # 如果只有一个文件，直接下载
            if len(results) == 1:
                st.download_button(
                    label="下载处理后的文件",
                    data=results[0][1],
                    file_name=results[0][0],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # 如果有多个文件，打包成 ZIP
            elif len(results) > 1:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for name, data in results:
                        zf.writestr(name, data)
                
                st.download_button(
                    label=f"📦 下载全部（{len(results)}个文件打包）",
                    data=zip_buffer.getvalue(),
                    file_name="处理后文件.zip",
                    mime="application/zip"
                )
            
            if results:
                st.balloons()
            else:
                st.error("所有文件处理失败")
