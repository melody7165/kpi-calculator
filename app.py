import streamlit as st
import openpyxl
from io import BytesIO

# 設定網頁標題與寬版排版
st.set_page_config(page_title="KPI 績效管理系統 v5", layout="wide")

# ==========================================
# 根據最新 Excel 檔案結構定義精確座標 (1-based)
# ==========================================
ROWS_1_TO_4 = [3, 4, 5, 6, 7, 8, 9, 10, 11]  # 評量指標 1~4
ROWS_5_TO_6 = [13, 14, 15, 16, 17, 18, 19]   # 加減分項 5~6
ROW_TEXT = 21                                # 加扣分事蹟 (列號)
ROW_HEADER = 2                               # 標題列 (員工姓名所在列)
COL_WEIGHT = 3                               # 分數佔比欄 (Column C)
COL_EMP_START = 4                            # 員工名單起始欄 (Column D, 即熊婷華)

def safe_float(val):
    """安全轉換數值，處理 Excel 中的 None 或特殊字串"""
    try:
        if val is None or str(val).strip() == "": return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0

# 側邊欄模式切換
st.sidebar.title("🛠️ KPI 管理系統")
app_mode = st.sidebar.radio("請選擇操作方式：", ["模式 A：合併兩份 Excel 檔", "模式 B：線上直接填表評分"])

# ---------------------------------------------------------
# 模式 A：兩份檔案自動合併
# ---------------------------------------------------------
if app_mode == "模式 A：合併兩份 Excel 檔":
    st.title("📂 模式 A：兩份 Excel 自動彙整")
    st.info("請上傳兩位組長的原始評分檔案。系統將自動計算平均分與累加分，並保留原有的公式樣式。")
    
    col1, col2 = st.columns(2)
    with col1: f1 = st.file_uploader("📂 上傳組長 A 檔案", type=["xlsx"], key="mA1")
    with col2: f2 = st.file_uploader("📂 上傳組長 B 檔案", type=["xlsx"], key="mA2")

    if f1 and f2:
        if st.button("🚀 開始彙整計算", use_container_width=True):
            wb1 = openpyxl.load_workbook(f1) # 保留樣式模板
            wb2 = openpyxl.load_workbook(f2, data_only=True)
            ws1, ws2 = wb1.active, wb2.active
            
            # 從第 4 欄 (D欄) 開始抓取有名字的人
            emp_cols = []
            for c in range(COL_EMP_START, ws1.max_column + 1):
                name = ws1.cell(ROW_HEADER, c).value
                if name: emp_cols.append(c)

            for c in emp_cols:
                # 1~4項：平均 (若兩方都有分則平均，單方有分則取該分)
                for r in ROWS_1_TO_4:
                    v1, v2 = safe_float(ws1.cell(r, c).value), safe_float(ws2.cell(r, c).value)
                    if v1 != 0 and v2 != 0: ws1.cell(r, c).value = (v1 + v2) / 2
                    elif v2 != 0: ws1.cell(r, c).value = v2
                # 5~6項：直接累加
                for r in ROWS_5_TO_6:
                    ws1.cell(r, c).value = safe_float(ws1.cell(r, c).value) + safe_float(ws2.cell(r, c).value)
                # 事蹟合併
                t1 = str(ws1.cell(ROW_TEXT, c).value or "").strip()
                t2 = str(ws2.cell(ROW_TEXT, c).value or "").strip()
                if t1 and t2: ws1.cell(ROW_TEXT, c).value = f"【組長A】:\n{t1}\n\n【組長B】:\n{t2}"
                elif t2: ws1.cell(ROW_TEXT, c).value = t2

            out = BytesIO(); wb1.save(out); out.seek(0)
            st.success("✅ 合併完成！")
            st.download_button("📥 下載彙整結果 Excel", out, "KPI_合併彙整結果.xlsx")

# ---------------------------------------------------------
# 模式 B：線上評分系統 (已校準 熊婷華 與 分數佔比)
# ---------------------------------------------------------
elif app_mode == "模式 B：線上直接填表評分":
    if "scores" not in st.session_state: st.session_state.scores = {}

    st.title("📝 模式 B：線上直接評分系統")
    st.sidebar.divider()
    temp_file = st.sidebar.file_uploader("📂 第一步：上傳空白 Excel 範本", type=["xlsx"], key="mB_temp")

    if temp_file:
        wb_temp = openpyxl.load_workbook(temp_file, data_only=True)
        ws_temp = wb_temp.active
        
        # 1. 抓取名單 (從 Column D 開始)
        employees = {}
        for c in range(COL_EMP_START, ws_temp.max_column + 1):
            name = ws_temp.cell(row=ROW_HEADER, column=c).value
            if name: employees[str(name).strip()] = c
        
        st.sidebar.write("👥 **偵測到的人員名單：**")
        st.sidebar.info("、".join(employees.keys()))
        
        # 2. 抓取指標與分數佔比 (Column C)
        item_info = {}
        curr_cat = ""
        for r in ROWS_1_TO_4 + ROWS_5_TO_6:
            cat = ws_temp.cell(r, 1).value
            if cat: curr_cat = str(cat).strip()
            desc = str(ws_temp.cell(r, 2).value or "").strip()
            weight = str(ws_temp.cell(r, COL_WEIGHT).value or "").strip() # 鎖定 Column C
            
            item_info[r] = {
                "title": f"{curr_cat} | {desc}" if curr_cat and curr_cat not in desc else desc,
                "weight": weight if weight else "無"
            }
        
        tab1, tab2 = st.tabs(["✍️ 開始評分", "📊 彙整下載"])
        
        with tab1:
            col_l, col_e = st.columns(2)
            with col_l: leader = st.text_input("👤 組長姓名：", placeholder="請輸入您的名字")
            with col_e: target_emp = st.selectbox("👥 評分對象：", ["-- 請選擇人員 --"] + list(employees.keys()))
            
            if leader and target_emp != "-- 請選擇人員 --":
                with st.form(f"form_{leader}_{target_emp}"):
                    st.write(f"### 🎯 為 **{target_emp}** 評分")
                    old_data = st.session_state.scores.get(leader, {}).get(target_emp, {})
                    new_scores = {}
                    
                    st.subheader("一、 評量指標 (1~4項)")
                    for r in ROWS_1_TO_4:
                        st.write(f"**{item_info[r]['title']}**")
                        st.markdown(f"<span style='color:#007BFF;'>📊 分數佔比/上限：{item_info[r]['weight']}</span>", unsafe_allow_html=True)
                        new_scores[r] = st.number_input("輸入得分", value=float(old_data.get(r, 0.0)), format="%.3f", key=f"v_{r}")
                        st.divider()
                        
                    st.subheader("二、 加減分項 (5~6項)")
                    for r in ROWS_5_TO_6:
                        st.write(f"**{item_info[r]['title']}**")
                        st.markdown(f"<span style='color:#FF4B4B;'>📌 加減分參考：{item_info[r]['weight']}</span>", unsafe_allow_html=True)
                        new_scores[r] = st.number_input("輸入得分", value=float(old_data.get(r, 0.0)), format="%.3f", key=f"v_{r}")
                        st.divider()
                        
                    st.subheader("三、 加扣分事蹟說明")
                    new_scores[ROW_TEXT] = st.text_area("請填寫具體事蹟", value=old_data.get(ROW_TEXT, ""), height=150)
                    
                    if st.form_submit_button("💾 儲存此員評分資料", use_container_width=True):
                        if leader not in st.session_state.scores: st.session_state.scores[leader] = {}
                        st.session_state.scores[leader][target_emp] = new_scores
                        st.success(f"✅ 已暫存 {leader} 對 {target_emp} 的評分。")

        with tab2:
            if st.session_state.scores:
                st.write("📋 **當前填寫進度：**")
                for l, d in st.session_state.scores.items(): st.write(f"- {l}：已完成 ({', '.join(d.keys())})")
                
                if st.button("🚀 執行 KPI 全員結算並產生 Excel", type="primary", use_container_width=True):
                    wb_out = openpyxl.load_workbook(template_file)
                    ws_out = wb_out.active
                    for emp, col_idx in employees.items():
                        v_ls = [l for l in st.session_state.scores if emp in st.session_state.scores[l]]
                        num = len(v_ls)
                        if num > 0:
                            for r in ROWS_1_TO_4: ws_out.cell(r, col_idx).value = sum(st.session_state.scores[l][emp].get(r, 0.0) for l in v_ls) / num
                            for r in ROWS_5_TO_6: ws_out.cell(r, col_idx).value = sum(st.session_state.scores[l][emp].get(r, 0.0) for l in v_ls)
                            ws_out.cell(ROW_TEXT, col_idx).value = "\n\n".join([f"【{l}】:\n{st.session_state.scores[l][emp].get(ROW_TEXT, '')}" for l in v_ls])
                    
                    out_b = BytesIO(); wb_out.save(out_b); out_b.seek(0)
                    st.success("🎉 結算完成！")
                    st.download_button("📥 下載最終 KPI 報表", out_b, "KPI_線上評分結算結果.xlsx", use_container_width=True)
            else:
                st.info("尚無任何評分紀錄。")
    else:
        st.info("👈 請先上傳 Excel 範本以開始。")
