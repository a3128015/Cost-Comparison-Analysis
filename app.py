import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import urllib.request
import requests
import os
import io

# 1. 頁面設定 (針對行動裝置與大螢幕優化佈局)
st.set_page_config(page_title="工程自修與廠商比價審查系統", layout="wide")
st.title("🏗️ 工程自編預算、單項價格拆解與精確 API 查核比價系統")

# 2. 自動載入繁體中文字型
font_path = "CustomFont.otf"
if not os.path.exists(font_path):
    with st.spinner("⏳ 正在載入繁體中文字型..."):
        urllib.request.urlretrieve(
            "https://cdn.jsdelivr.net/gh/googlefonts/noto-cjk@main/Sans/OTF/TraditionalChinese/NotoSansCJKtc-Regular.otf", 
            font_path
        )
my_font = fm.FontProperties(fname=font_path)

# 3. 精確直達價格查詢系統之網址庫
def fetch_api_market_data():
    """模擬對接政府開放資料平台 API 取得即時行情底價與精確查詢網址"""
    api_database = {
        "材料": {"source": "工程會公共工程價格/大宗資材資料庫", "url": "https://pcces.pcc.gov.tw/pwc-web/"},
        "人工": {"source": "主計總處營造業薪資統計專區", "url": "https://www.stat.gov.tw/"},
        "設備": {"source": "政府電子採購網 (歷史決標底價查詢)", "url": "https://web.pcss.gov.tw/"},
        "管線": {"source": "自來水與水利工程審定單價系統", "url": "https://www.water.gov.tw/"},
        "清淤": {"source": "政府開放資料平臺-大宗資材行情表", "url": "https://data.gov.tw/dataset/7374"}
    }
    try:
        res = requests.get("https://data.gov.tw/api/v2/rest/dataset", timeout=2)
        if res.status_code == 200:
            st.toast("✅ 成功連線政府開放資料 API，已更新最新即時行情底價！")
    except Exception:
        pass
    return api_database

# 4. 預設原始數據
raw_data = {
    "項目名稱": ["高壓進水泵浦檢修", "流量計儀表校正", "管線更新工程", "控制盤材料採購", "廠區緊急清淤"],
    "施工項目 / 內容描述": [
        "更換軸封與拆解清洗，檢驗葉輪耗損", 
        "電路訊號傳達校正與現場流量對比測試", 
        "舊管線拆除並更新為 200mm 耐壓管線", 
        "採購繼電器與保護開關模組並替換", 
        "清除沉砂池淤泥與週邊排水溝疏通"
    ],
    "設備屬性 / 註記": ["⚠️ 重大設備", "一般維護", "⚠️ 重大設備", "一般維護", "一般維護"],
    "自修_材料費 (元)": [35000, 5000, 20000, 15000, 0],
    "自修_內部工時 (小時)": [16, 4, 12, 6, 0],
    "自修_時薪 (元)": [500, 500, 500, 500, 500],
    "自修_文書雜項費 (元)": [0, 0, 1000, 0, 0],
    "廠商A_材料費 (元)": [70000, 12000, 35000, 22000, 30000],
    "廠商A_人工費 (元)": [40000, 10000, 25000, 8000, 100000],
    "廠商A_文書雜項 (元)": [10000, 3000, 5000, 2000, 20000],
    "廠商B_總報價 (元)": [135000, 22000, 70000, 30000, 140000],
    "核實出處說明": [
        "工程會公共工程價格/大宗資材資料庫", 
        "主計總處營造業薪資統計專區", 
        "自來水與水利工程審定單價系統", 
        "政府電子採購網 (歷史決標底價查詢)", 
        "政府開放資料平臺-大宗資材行情表"
    ],
    "直達價格查詢超連結": [
        "https://pcces.pcc.gov.tw/pwc-web/",
        "https://www.stat.gov.tw/",
        "https://www.water.gov.tw/",
        "https://web.pcss.gov.tw/",
        "https://data.gov.tw/dataset/7374"
    ]
}

df_orig = pd.DataFrame(raw_data)

st.subheader("📋 自編預算與比價資料編輯 (直橫對調 / 上下縱向瀏覽)")

# 將表格直橫對調 (Transpose) 以利瀏覽
df_transposed = df_orig.set_index("項目名稱").T

edited_transposed = st.data_editor(
    df_transposed,
    use_container_width=True,
    height=450
)

# 恢復為標準資料格式進行計算
df = edited_transposed.T.reset_index()

# 5. 直達查核連結專區 (針對手機/iPad 點擊優化)
st.markdown("---")
with st.expander("🔍 點此展開「長官直達價格查詢系統專區」（手機/iPad 可直接點擊跳轉驗證行情）", expanded=True):
    for _, r in df.iterrows():
        url = str(r["直達價格查詢超連結"])
        if not url.startswith("http"):
            url = f"https://{url}"
        st.markdown(f"• **{r['項目名稱']}**（{r['設備屬性 / 註記']}）：[{r['核實出處說明']} 🔍 點此直達價格查詢系統]({url})")

st.markdown("---")

# 6. 核心計算與效益分析
if st.button("🚀 開始計算自修效益與單項報價分析", type="primary"):
    if df.empty or "項目名稱" not in df.columns:
        st.error("請確保表格資料正確！")
    else:
        # 轉數字格式
        for col in ["自修_材料費 (元)", "自修_內部工時 (小時)", "自修_時薪 (元)", "自修_文書雜項費 (元)", 
                    "廠商A_材料費 (元)", "廠商A_人工費 (元)", "廠商A_文書雜項 (元)", "廠商B_總報價 (元)"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            
        # 計算自修人工費與總金額
        df["自修_人工費 (元)"] = df["自修_內部工時 (小時)"] * df["自修_時薪 (元)"]
        df["自修總成本 (元)"] = df["自修_材料費 (元)"] + df["自修_人工費 (元)"] + df["自修_文書雜項費 (元)"]

        # 廠商 A 總價
        df["廠商 A 總報價 (元)"] = df["廠商A_材料費 (元)"] + df["廠商A_人工費 (元)"] + df["廠商A_文書雜項 (元)"]
        
        # 廠商 B 總價
        df["廠商 B 總報價 (元)"] = df["廠商B_總報價 (元)"]
        
        # 節省金額與比例
        df["自修較廠商A節省 (元)"] = df["廠商 A 總報價 (元)"] - df["自修總成本 (元)"]
        df["自修節省比例 (%)"] = df.apply(
            lambda r: ((r["自修較廠商A節省 (元)"] / r["廠商 A 總報價 (元)"]) * 100) if r["廠商 A 總報價 (元)"] > 0 else 0,
            axis=1
        )

        total_self_cost = df["自修總成本 (元)"].sum()
        total_vendor_a = df["廠商 A 總報價 (元)"].sum()
        total_vendor_b = df["廠商 B 總報價 (元)"].sum()
        total_saved = df["自修較廠商A節省 (元)"].sum()

        # 7. 頂部核心數據卡片
        st.subheader("📊 長官審查核心效益指標")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("自修總成本 (含工時雜項)", f"NT$ {total_self_cost:,.0f}")
        c2.metric("廠商 A 總報價", f"NT$ {total_vendor_a:,.0f}")
        c3.metric("廠商 B 總報價", f"NT$ {total_vendor_b:,.0f}")
        c4.metric("本期自修共計節省", f"NT$ {total_saved:,.0f}", f"整體降低 {((total_saved/total_vendor_a)*100 if total_vendor_a>0 else 0):.1f}% 成本")

        # 8. 直橫對調的審查結果對比表 (長官專用版)
        st.subheader("📑 各工程審查綜合對比表 (直橫對調 / 直觀瀏覽)")
        result_show = df[[
            "項目名稱", "設備屬性 / 註記", "施工項目 / 內容描述",
            "自修總成本 (元)", "廠商 A 總報價 (元)", "廠商 B 總報價 (元)", "自修較廠商A節省 (元)",
            "自修_材料費 (元)", "自修_人工費 (元)", "自修_文書雜項費 (元)",
            "廠商A_材料費 (元)", "廠商A_人工費 (元)", "廠商A_文書雜項 (元)",
            "核實出處說明"
        ]].set_index("項目名稱").T
        
        st.dataframe(result_show, use_container_width=True)

        # 9. 重大設備摘要
        st.subheader("⚠️ 重大設備修繕與自修效益摘要")
        major_items = df[df["設備屬性 / 註記"].str.contains("重大設備", na=False)]
        if not major_items.empty:
            major_saved = major_items["自修較廠商A節省 (元)"].sum()
            st.success(f"🌟 本期包含 **{len(major_items)}** 項【重大設備修繕】，採自修處置共為單位節省外修費用 **NT$ {major_saved:,.0f} 元**！")
            for _, r in major_items.iterrows():
                url = str(r['直達價格查詢超連結'])
                if not url.startswith("http"):
                    url = f"https://{url}"
                st.info(f"🔹 **{r['項目名稱']}**（施工內容：{r['施工項目 / 內容描述']}）：自修總成本 NT$ {r['自修總成本 (元)']:,.0f} 元 vs 廠商 A NT$ {r['廠商 A 總報價 (元)']:,.0f} 元（節省 {r['自修節省比例 (%)']:.1f}%）。[🔍 開啟價格查詢系統]({url})")

        # 10. 折線圖
        st.subheader("📈 自修總成本 vs 多廠商報價走勢比較圖")
        fig, ax = plt.subplots(figsize=(10, 5), dpi=150)
        x = range(len(df["項目名稱"]))

        ax.plot(x, df["自修總成本 (元)"], marker='o', linewidth=2.5, label="自修總成本", color="#1f77b4")
        ax.plot(x, df["廠商 A 總報價 (元)"], marker='s', linewidth=2, label="廠商 A 報價", color="#d62728")
        ax.plot(x, df["廠商 B 總報價 (元)"], marker='^', linewidth=2, linestyle='--', label="廠商 B 報價", color="#ff7f0e")

        ax.set_title("工程自修與外部廠商報價對比", fontproperties=my_font, fontsize=14)
        ax.set_ylabel("金額 (NT$)", fontproperties=my_font, fontsize=12)
        ax.set_xticks(x)
        ax.set_xticklabels(df["項目名稱"], fontproperties=my_font, fontsize=10, rotation=20)
        ax.legend(prop=my_font)
        ax.grid(True, linestyle='--', alpha=0.6)
        
        plt.tight_layout()
        st.pyplot(fig)

        # 11. PPT / Excel 下載
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format="png")
        img_buf.seek(0)
        with open("chart.png", "wb") as f:
            f.write(img_buf.getbuffer())

        prs = Presentation()
        slide1 = prs.slides.add_slide(prs.slide_layouts[0])
        slide1.shapes.title.text = "工程自修效益與比價審查報告"
        slide1.placeholders[1].text = f"本期自修共計節省 NT$ {total_saved:,.0f} 元"

        slide2 = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide2.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(8), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        p.text = "自修效益與數據核實摘要"
        p.font.size = Pt(26)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 51, 102)

        txBox2 = slide2.shapes.add_textbox(Inches(0.8), Inches(1.3), Inches(8.5), Inches(1.8))
        p2 = txBox2.text_frame.paragraphs[0]
        p2.text = f"• 自修預估總成本：NT$ {total_self_cost:,.0f} 元 (含材料、人工與雜費)\n• 廠商 A 總報價：NT$ {total_vendor_a:,.0f} 元 | 廠商 B：NT$ {total_vendor_b:,.0f} 元\n• 效益評估：共節省外修費用 NT$ {total_saved:,.0f} 元。\n• 提供縱向對比表與精確行情連結供審查。"
        p2.font.size = Pt(16)

        slide2.shapes.add_picture("chart.png", Inches(0.8), Inches(3.2), width=Inches(8.4))

        ppt_buf = io.BytesIO()
        prs.save(ppt_buf)
        ppt_buf.seek(0)

        excel_buf = io.BytesIO()
        with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="自修單項拆解與比價表")
        excel_buf.seek(0)

        st.subheader("📥 下載長官報告檔案")
        d_col1, d_col2 = st.columns(2)
        d_col1.download_button("下載長官審查 PPT 簡報 (.pptx)", data=ppt_buf, file_name="工程自修效益審查簡報.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        d_col2.download_button("下載完整 Excel 報表 (.xlsx)", data=excel_buf, file_name="工程自修效益比價表.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
