import streamlit as st
from processor import process_excel

st.set_page_config(page_title="Excel 处理器", layout="centered")

st.title("📊 Excel 数据处理工具")
st.markdown("上传 `.xls` 或 `.xlsx` 文件，自动计算排放量、平均值和最大值。")

uploaded_file = st.file_uploader("选择文件", type=["xls", "xlsx"])

if uploaded_file is not None:
    file_name = uploaded_file.name
    
    with st.spinner("正在处理..."):
        try:
            result_bytes = process_excel(uploaded_file.getvalue(), file_name)
            
            st.success("处理完成！")
            st.download_button(
                label="📥 下载处理后的文件",
                data=result_bytes,
                file_name=file_name.rsplit(".", 1)[0] + "_处理后.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except ValueError as e:
            st.warning(str(e))
        except Exception as e:
            st.error(f"处理失败：{str(e)}")
