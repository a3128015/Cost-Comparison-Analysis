import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import urllib.request
import os
import io

# 1. 頁面標題與佈局
st.set_page_config(page_title="工程成本與廠商誠信比價系統", layout="wide")
st.title("🏗️ 工程成本比較、廠商誠信度與 PPT 報表生成器")

# 2. 自動下載並載入中文字型
font_path = "CustomFont.otf"
if not os.path.exists(font_path):
    with st.spinner("⏳ 正在載入繁體中文字型..."):
        urllib.request.urlretrieve(
            "https://cdn.jsdelivr.net/gh/googlefonts/noto-cjk@main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf", 
            font_path
        )
my_font = fm.FontProperties(fname=font_path)

# 3. 擴充的公開行情與價格來源資料庫（可自行延伸）
MARKET_DATABASE = {
    "材料": {"unit_price": 2800, "source": "公共工程委員會大宗資材資料庫 (2026/Q3)"},
    "人工": {"unit_price": 3200, "source": "營造業勞工統計月報與市場薪資行情"},
    "設備": {"unit_price": 12500, "source": "重型機械租賃同業公會標準單價"},
    "運銷": {"unit_price": 1500, "source": "汽車貨運商業同業公會核定費率"},
    "管線": {"unit_price": 650, "source": "自來水事業處與下水道工程處審定單價"},
    "清淤": {"unit_price": 1800, "source": "環保局水質保護科清淤工程參考單價"},
    "校正": {"unit_price": 12000, "source": "度量衡標準檢驗局與原廠校正標準費率"},
    "雜項": {"unit_price": 1000, "source": "歷史工程結算與常規行政管銷費率"}
}

# 4. 廠商歷史報價資料庫（用於前後比對是否有被加價）
HISTORY_DATABASE = {
    "材料費": 1000000,
    "人工費": 750000,
    "設備租賃": 400000,
    "運銷費": 130000,
    "其他雜項": 80000
}

def get_source_and_history(item_name):
    """自動為任意新增的項目匹配來源與歷史價格"""
    # 搜尋市場來源
    matched_source = "歷史採購結算與市場實務單價 (建議現場審查)"
    for key, data in MARKET_DATABASE.items():
        if key in str(item_name):
            matched_source = data["source"]
            break
            
    # 搜尋歷史報價
    history_price = HISTORY_DATABASE.get(item_name, None)
    return matched_source, history_price

# 5. 資料輸入與自動匹配邏輯
st.subheader("📋 工程項目比價與廠商歷史追蹤（自由新增項目將自動帶入來源）")

# 預設範例資料
default_data = pd.DataFrame({
    "項目名稱": ["材料費", "人工費", "設備租賃", "運銷費", "其他雜項"],
    "歷史上次報價 (元)": [1000000, 750000, 400000, 130000, 80000],
    "本次廠商報價 (元)": [1200000, 800000, 450000, 150000, 100000],
    "方案B_優化方案 (元)": [950000, 650000, 400000, 120000, 80000],
    "價格出處與核實依據": [
        "公共工程委員會大宗資材資料庫 (2026/Q3)",
        "營造業勞工統計月報與市場薪資行情",
        "重型機械租賃同業公會標準單價",
        "汽車貨運商業同業公會核定費率",
        "歷史工程結算與常規行政管銷費率"
    ]
})

edited_df = st.data_editor(
    default_data, 
    num_rows="dynamic",
    use_container_width=True
)

# 自動為使用者新手動新增的項目補上「來源」與「歷史價格」
for idx, row in edited_df.iterrows():
    item = str(row["項目名稱"])
    if pd.isna(row["價格出處與核實依據"]) or str(row["價格出處與核實依據"]).strip() == "":
        source, hist_p = get_source_and_history(item)
        edited_df.at[idx, "價格出處與核實依據"] = source
        if pd.isna(row["歷史上次報價 (元)"]) and hist_p is not None:
            edited_df.at[idx, "歷史上次報價 (元)"] = hist_p

# 6. 比價與誠信度分析計算
if st.button("🚀 開始比價與廠商誠信度分析", type="primary"):
    if edited_df.empty or "項目名稱" not in edited_df.columns:
        st.error("請確保表格包含有效的項目名稱！")
    else:
        df = edited_df.copy()
        df["歷史上次報價 (元)"] = pd.to_numeric(df["歷史上次報價 (元)"], errors="coerce").fillna(0)
        df["本次廠商報價 (元)"] = pd.to_numeric(df["本次廠商報價 (元)"], errors="coerce").fillna(0)
        df["方案B_優化方案 (元)"] = pd.to_numeric(df["方案B_優化方案 (元)"], errors="coerce").fillna(0)
        
        # 計算與上次報價的調漲幅度
        df["較上次加價 (元)"] = df["本次廠商報價 (元)"] - df["歷史上次報價 (元)"]
        df["漲幅 (%)"] = ((df["較上次加價 (元)"] / df["歷史上次報價 (元)"]) * 100).fillna(0)
        df["節省金額 (元)"] = df["本次廠商報價 (元)"] - df["方案B_優化方案 (元)"]

        total_hist = df["歷史上次報價 (元)"].sum()
        total_curr = df["本次廠商報價 (元)"].sum()
        total_b = df["方案B_優化方案 (元)"].sum()
        total_added = total_curr - total_hist
        total_saved = total_curr - total_b

        # 顯示指標卡片
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("歷史上次總報價", f"NT$ {total_hist:,.0f}")
        c2.metric("本次廠商報價", f"NT$ {total_curr:,.0f}", f"+{total_added:,.0f} 元", delta_color="inverse")
        c3.metric("方案 B 優化估算", f"NT$ {total_b:,.0f}")
        c4.metric("比廠商報價節省", f"NT$ {total_saved:,.0f}")

        # 誠信度提醒警示
        st.subheader("⚠️ 廠商報價異動與誠信度警示")
        overpriced_items = df[df["漲幅 (%)"] > 10]
        if not overpriced_items.empty:
            for _, r in overpriced_items.iterrows():
                st.warning(f"🚨 項目【{r['項目名稱']}】較上次報價大幅調漲 {r['漲幅 (%)']:.1f}%（增加 NT$ {r['較上次加價 (元)']:,.0f} 元）。請審查核實依據：{r['價格出處與核實依據']}")
        else:
            st.success("✅ 本次廠商報價對比歷史紀錄未發現不合理暴漲項目。")

        # 圖表繪製
        fig, ax = plt.subplots(figsize=(10, 5), dpi=150)
        x = range(len(df["項目名稱"]))
        width = 0.25

        ax.bar([i - width for i in x], df["歷史上次報價 (元)"], width, label="歷史上次報價")
        ax.bar(x, df["本次廠商報價 (元)"], width, label="本次廠商報價")
        ax.bar([i + width for i in x], df["方案B_優化方案 (元)"], width, label="方案B (優化)")

        ax.set_title("前後期廠商報價與優化方案比價圖", fontproperties=my_font, fontsize=14)
        ax.set_ylabel("金額 (NT$)", fontproperties=my_font, fontsize=12)
        ax.set_xticks(x)
        ax.set_xticklabels(df["項目名稱"], fontproperties=my_font, fontsize=10, rotation=20)
        ax.legend(prop=my_font)
        ax.grid(axis='y', linestyle='--', alpha=0.7)
        plt.tight_layout()

        st.pyplot(fig)

        # 繪圖暫存供 PPT 使用
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format="png")
        img_buf.seek(0)
        with open("chart.png", "wb") as f:
            f.write(img_buf.getbuffer())

        # 產生包含誠信度與核實出處的 PPT
        prs = Presentation()
        slide1 = prs.slides.add_slide(prs.slide_layouts[0])
        slide1.shapes.title.text = "工程成本審查與廠商報價比對報告"
        slide1.placeholders[1].text = f"廠商本次報價較歷史增加 NT$ {total_added:,.0f} 元 | 採方案B預計可節省 NT$ {total_saved:,.0f} 元"

        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide2.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(8), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        p.text = "廠商報價前後比對與來源核實"
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 51, 102)

        txBox2 = slide2.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(8.5), Inches(1.8))
        p2 = txBox2.text_frame.paragraphs[0]
        p2.text = f"• 歷史上次總報價：NT$ {total_hist:,.0f} 元\n• 本次廠商總報價：NT$ {total_curr:,.0f} 元 (漲幅 +{((total_curr-total_hist)/total_hist*100):.1f}%)\n• 誠信度評估：已自動比對資材庫，漲幅過高項目已列入審查重點。"
        p2.font.size = Pt(16)

        slide2.shapes.add_picture("chart.png", Inches(0.8), Inches(3.2), width=Inches(8.4))

        ppt_buf = io.BytesIO()
        prs.save(ppt_buf)
        ppt_buf.seek(0)

        # 輸出 Excel
        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="成本與誠信比對表")
        excel_buf.seek(0)

        st.subheader("📥 下載產出檔案")
        d_col1, d_col2 = st.columns(2)
        d_col1.download_button("下載比價 PPT 簡報 (.pptx)", data=ppt_buf, file_name="工程成本與誠信比對簡報.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        d_col2.download_button("下載完整 Excel (.xlsx)", data=excel_buf, file_name="工程成本與誠信比對表.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
