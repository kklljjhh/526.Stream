import streamlit as st
from processor import process_excel

st.set_page_config(page_title="Excel 智能清洗工具", layout="centered")

st.title("📊 Excel 智能清洗工具")
st.markdown("上传 .xls / .xlsx 文件，自动处理排放量计算、标记颜色、统计汇总")

uploaded = st.file_uploader("上传 Excel 文件", type=["xlsx", "xls"])

if uploaded:
    st.success(f"已接收：{uploaded.name} ({round(len(uploaded.getvalue())/1024, 1)} KB)")
    
    if st.button("开始处理", type="primary"):
        with st.spinner("正在处理..."):
            try:
                result_bytes = process_excel(uploaded.getvalue(), uploaded.name)
                
                st.download_button(
                    label="下载处理后的文件",
                    data=result_bytes,
                    file_name=f"{uploaded.name.rsplit('.', 1)[0]}_处理后.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.balloons()
                
            except ValueError as e:
                if "实时数据" in str(e):
                    st.warning(f"⚠️ {e}")
                else:
                    st.error(f"❌ {e}")
            except Exception as e:
                st.error(f"处理出错：{e}")
