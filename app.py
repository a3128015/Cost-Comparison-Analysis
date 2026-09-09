import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
import io

st.set_page_config(page_title="工程成本比價與簡報生成器", layout="wide")

st.title("🏗️ 工程成本比較與自動簡報生成器")
st.write("請在下方輸入或上傳工程比價數據，系統將自動生成分析圖表與 PowerPoint 簡報。")

# 1. 數據輸入區
st.header("1. 輸入工程項目與報價")

default_data = {
    "項目名稱": ["混凝土工程", "鋼筋工程", "模板工程", "開挖工程"],
    "廠商A (元)": [150000, 280000, 120000, 90000],
    "廠商B (元)": [140000, 295000, 115000, 95000],
    "廠商C (元)": [160000, 270000, 125000, 88000]
}

df = st.data_editor(pd.DataFrame(default_data), num_rows="dynamic")

# 2. 數據分析與圖表
if not df.empty:
    st.header("2. 成本比較分析")
    
    # 計算總價
    df_sum = df.drop(columns=["項目名稱"]).sum().reset_index()
    df_sum.columns = ["廠商", "總成本 (元)"]
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("廠商總報價比較")
        st.dataframe(df_sum, hide_index=True)
        
    with col2:
        st.subheader("報價柱狀圖")
        fig, ax = plt.subplots()
        ax.bar(df_sum["廠商"], df_sum["總成本 (元)"], color=['#1f77b4', '#ff7f0e', '#2ca02c'])
        ax.set_ylabel("金額 (元)")
        st.pyplot(fig)

    # 3. PPT 生成與下載
    st.header("3. 自動生成 PPT 簡報")
    
    if st.button("🚀 產生 PowerPoint 簡報"):
        prs = Presentation()
        slide_layout = prs.slide_layouts[6] # 空白簡報頁
        slide = prs.slides.add_slide(slide_layout)
        
        # 新增標題
        txBox = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = "工程成本比價分析報告"
        p.font.size = Pt(32)
        p.font.bold = True
        
        # 儲存圖表並放入 PPT
        img_buf = io.BytesIO()
        fig.savefig(img_buf, format='png', bbox_inches='tight')
        img_buf.seek(0)
        slide.shapes.add_picture(img_buf, Inches(1), Inches(1.8), width=Inches(6.5))
        
        # 下載 PPT
        ppt_buf = io.BytesIO()
        prs.save(ppt_buf)
        ppt_buf.seek(0)
        
        st.download_button(
            label="📥 下載 PPT 簡報",
            data=ppt_buf,
            file_name="Cost_Comparison_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
