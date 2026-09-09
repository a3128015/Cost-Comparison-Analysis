import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
import io
import pdfplumber

st.set_page_config(page_title="工程成本比價與簡報生成器", layout="wide")

st.title("🏗️ 工程成本比較與自動簡報生成器")
st.write("您可以選擇<b>上傳 Excel / CSV / PDF 報價單</b>，或直接在下方表格中<b>雙擊儲存格</b>自由修改與新增廠商/項目數據：")

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
                    cleaned_row = [str(cell).strip() if cell is not None else "" for cell in row]
                    if any(cleaned_row):
                        extracted_data.append(cleaned_row)
    
    if extracted_data:
        header = extracted_data[0]
        rows = extracted_data[1:]
        df_pdf = pd.DataFrame(rows, columns=header)
        return df_pdf
    else:
        return None

# 初始化或載入資料
if "df_data" not in st.session_state:
    st.session_state.df_data = pd.DataFrame(default_data)

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith('.csv'):
            st.session_state.df_data = pd.read_csv(uploaded_file)
            st.success("CSV 檔案上傳成功！")
        elif uploaded_file.name.endswith('.pdf'):
            parsed_df = parse_pdf(io.BytesIO(uploaded_file.read()))
            if parsed_df is not None and not parsed_df.empty:
                st.session_state.df_data = parsed_df
                st.success("PDF 報價單辨識成功！請在下方檢查與修正欄位資訊。")
            else:
                st.warning("PDF 中未偵測到標準表格，套用預設範本。")
        else:
            st.session_state.df_data = pd.read_excel(uploaded_file)
            st.success("Excel 檔案上傳成功！")
    except Exception as e:
        st.error(f"檔案讀取失敗：{e}")

# 工具列：手動新增廠商欄位
col_btn1, col_btn2 = st.columns([1, 4])
with col_btn1:
    new_vendor_name = st.text_input("新增廠商名稱", value="", placeholder="例如：廠商D")
with col_btn2:
    st.write("")
    st.write("")
    if st.button("➕ 新增該廠商欄位"):
        if new_vendor_name and new_vendor_name not in st.session_state.df_data.columns:
            st.session_state.df_data[new_vendor_name] = 0
            st.rerun()

st.subheader("💡 互動式資料編輯表（按兩下儲存格即可修改，最下方可直接新增項目列）")

# 編輯器區塊
edited_df = st.data_editor(
    st.session_state.df_data,
    num_rows="dynamic",
    use_container_width=True,
    key="editor"
)

# 2. 數據分析與圖表
if not edited_df.empty and len(edited_df.columns) > 1:
    st.header("2. 成本比較分析")
    
    item_col = edited_df.columns[0]
    vendor_cols = edited_df.columns[1:]
    
    # 資料清理：將「元」、「,」等符號移除並轉為數值
    df_clean = edited_df[vendor_cols].copy()
    for col in vendor_cols:
        df_clean[col] = df_clean[col].astype(str).str.replace('元', '').str.replace(',', '').str.strip()
        df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce').fillna(0)
    
    df_sum = df_clean.sum().reset_index()
    df_sum.columns = ["廠商", "總成本 (元)"]
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("廠商總報價比較")
        st.dataframe(df_sum, hide_index=True, use_container_width=True)
        
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
        
        txBox = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        p.text = "工程成本比價分析報告"
        p.font.size = Pt(32)
        p.font.bold = True
        
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
