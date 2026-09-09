import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import re
import pdfplumber

st.set_page_config(page_title="工程成本比價與簡報生成器", layout="wide")

st.title("🏗️ 工程成本比較與自動簡報生成器")
st.write("您可以選擇<b>上傳 Excel / CSV / PDF 報價單</b>，或直接在下方表格中輸入與修改數據：")

# 1. 檔案上傳與數據輸入區
st.header("1. 上傳或輸入工程報價資料")

uploaded_file = st.file_uploader("選擇您的 Excel、CSV 或 PDF 報價單檔案", type=["xlsx", "xls", "csv", "pdf"])

default_data = {
    "項目名稱": ["混凝土工程", "鋼筋工程", "模板工程", "開挖工程"],
    "廠商A (元)": [150000, 280000, 120000, 90000],
    "廠商B (元)": [140000, 295000, 115000, 95000],
    "廠商C (元)": [160000, 270000, 125000, 88000]
}

def parse_pdf(file_bytes):
    """解析 PDF 檔中的表格與報價數據"""
    extracted_data = []
    with pdfplumber.open(file_bytes) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    # 過濾掉空列
                    cleaned_row = [str(cell).strip() if cell is not None else "" for cell in row]
                    if any(cleaned_row):
                        extracted_data.append(cleaned_row)
    
    if extracted_data:
        # 將第一列設為表頭，其餘為資料
        header = extracted_data[0]
        rows = extracted_data[1:]
        df_pdf = pd.DataFrame(rows, columns=header)
        return df_pdf
    else:
        return None

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            df_input = pd.read_csv(uploaded_file)
            st.success("CSV 檔案上傳成功！")
        elif uploaded_file.name.endswith('.pdf'):
            parsed_df = parse_pdf(io.BytesIO(uploaded_file.read()))
            if parsed_df is not None and not parsed_df.empty:
                df_input = parsed_df
                st.success("PDF 報價單辨識成功！請在下方檢查與修正欄位資訊。")
            else:
                st.warning("PDF 中未偵測到標準表格，已套用預設範本，請手動確認資料。")
                df_input = pd.DataFrame(default_data)
        else:
            df_input = pd.read_excel(uploaded_file)
            st.success("Excel 檔案上傳成功！")
    except Exception as e:
        st.error(f"檔案讀取或解析失敗：{e}")
        df_input = pd.DataFrame(default_data)
else:
    df_input = pd.DataFrame(default_data)

st.subheader("數據預覽與編輯區")
df = st.data_editor(df_input, num_rows="dynamic")

# 2. 數據分析與圖表
if not df.empty and len(df.columns) > 1:
    st.header("2. 成本比較分析")
    
    item_col = df.columns[0]
    vendor_cols = df.columns[1:]
    
    # 強制將數值欄位轉為數字以利統計計算
    df_sum = df[vendor_cols].apply(pd.to_numeric, errors='coerce').sum().reset_index()
    df_sum.columns = ["廠商", "總成本 (元)"]
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("廠商總報價比較")
        st.dataframe(df_sum, hide_index=True)
        
    with col2:
        st.subheader("報價柱狀圖")
        fig, ax = plt.subplots()
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b']
        ax.bar(df_sum["廠商"], df_sum["總成本 (元)"], color=colors[:len(df_sum)])
        ax.set_ylabel("金額 (元)")
        plt.xticks(rotation=15)
        st.pyplot(fig)

    # 3. PPT 生成與下載
    st.header("3. 自動生成 PPT 簡報")
    
    if st.button("🚀 產生 PowerPoint 簡報"):
        prs = Presentation()
        slide_layout = prs.slide_layouts[6]
        slide = prs.slides.add_slide(slide_layout)
        
        # 標題
        txBox = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = "工程成本比價分析報告"
        p.font.size = Pt(32)
        p.font.bold = True
        
        # 將圖表寫入 PPT
        img_buf = io.BytesIO()
        fig.savefig(img_buf, format='png', bbox_inches='tight')
        img_buf.seek(0)
        slide.shapes.add_picture(img_buf, Inches(1), Inches(1.8), width=Inches(6.5))
        
        ppt_buf = io.BytesIO()
        prs.save(ppt_buf)
        ppt_buf.seek(0)
        
        st.download_button(
            label="📥 下載 PPT 簡報",
            data=ppt_buf,
            file_name="Cost_Comparison_Report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
