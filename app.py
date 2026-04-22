import streamlit as st
import openpyxl
from io import BytesIO

# 設定網頁標題與寬版排版
st.set_page_config(page_title="KPI 績效管理系統 v3", layout="wide")

# 根據 26_Q1-3_個人績效評量_空.xlsx 定義精確位置
ROWS_1_TO_4 = [3, 4, 5, 6, 7, 8, 9, 10, 11]  # 平均項 (Excel 列號)
ROWS_5_TO_6 = [13, 14, 15, 16, 17, 18, 19]   # 累加項
ROW_TEXT = 21                                # 事蹟說明
NAME_ROW = 2                                 # 人員姓名所在列 (Excel Row 2)
WEIGHT_COL = 3                               # 分數佔比所在欄 (Column C)
START_DATA_COL = 4                           # 人員分數起始欄 (Column D, 即熊婷華開始)

def safe_float(val):
    try:
        if val is None or str(val).strip() == "": return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0

# 側邊欄：模式切換
st.sidebar.title("🛠️ KPI 系統選單")
app_mode = st.sidebar.radio("請選擇操作方式：", ["模式 A：合併兩份 Excel 檔", "模式 B：線上直接填表評分"])

# ---------------------------------------------------------
# 模式 A：兩份檔案自動合併
# ---------------------------------------------------------
if app_mode == "模式 A：合併兩份 Excel 檔":
    st.title("📂 模式 A：兩份 Excel 自動合併")
    st.info("請分別上傳兩位組長的評分檔案，系統將依邏輯合併分數並保留原格式。")
    
    col1, col2 = st.columns(2)
    with col1: f1 = st.file_uploader("📂 上傳【組長 A】評分檔", type=["xlsx"], key="mA_f1")
    with col2: f2 = st.file_uploader("📂 上傳【組長 B】評分檔", type=["xlsx"], key="mA_f2")

    if f1 and f2:
        if st.button("🚀 執行合併計算", use_container_width=True):
            wb1 = openpyxl.load_workbook(f1) 
            wb2 = openpyxl.load_workbook(f2, data_only=True)
            ws1, ws2 = wb1.active, wb2.active
            
            # 偵測人員欄位 (從 Column D 開始)
            emp_cols = [c for c in range(START_DATA_COL, ws1.max_column + 1) if ws1.cell(NAME_ROW, c).value]
            
            for c in emp_cols:
                # 1~4項 平均
                for r in ROWS_1_TO_4:
                    v1, v2 = safe_float(ws1.cell(r, c).value), safe_float(ws2.cell(r, c).value)
                    if v1 != 0 and v2 != 0: ws1.cell(r, c).value = (v1 + v2) / 2
                    elif v2 != 0: ws1.cell(r, c).value = v2
                # 5~6項 累加
                for r in ROWS_5_TO_6:
                    ws1.cell(r, c).value = safe_float(ws1.cell(r, c).value) + safe_float(ws2.cell(r, c).value)
                # 文字合併
                t1 = str(ws1.cell(ROW_TEXT, c).value or "").strip()
                t2 = str(ws2.cell(ROW_TEXT, c).value or "").strip()
                if t1 and t2: ws1.cell(ROW_TEXT, c).value = f"【組長A】:\n{t1}\n\n【組長B】:\n{t2}"
                elif t2: ws1.cell(ROW_TEXT, c).value = t2

            out = BytesIO(); wb1.save(out); out.seek(0)
            st.success("✅ 合併完成！")
            st.download_button("📥 下載合併結果 Excel", out, "KPI_合併結果.xlsx")

# ---------------------------------------------------------
# 模式 B：線上評分系統 (針對新版 Excel 優化)
# ---------------------------------------------------------
elif app_mode == "模式 B：線上直接填表評分":
    if "scores" not in st.session_state: st.session_state.scores = {}

    st.title("📝 模式 B：線上直接填表評分")
    st.sidebar.divider()
    template_file = st.sidebar.file_uploader("📂 第一步：上傳空白 Excel 範本", type=["xlsx"], key="mB_temp")

    if template_file:
        wb_temp = openpyxl.load_workbook(template_file, data_only=True)
        ws_temp = wb_temp.active
        
        # 1. 抓取員工名單 (精確從 D 欄開始，確保 熊婷華 在內)
        employees = {}
        for c in range(START_DATA_COL, ws_temp.max_column + 1):
            name = ws_temp.cell(row=NAME_ROW, column=c).value
            if name: employees[str(name).strip()] = c
        
        st.sidebar.write("✅ **偵測到的人員：**")
        st.sidebar.code(", ".join(employees.keys())) # 讓使用者確認名單
        
        # 2. 抓取評分項目與佔比 (處理合併單元格標題)
        item_info = {}
        current_cat = ""
        for r in ROWS_1_TO_4 + ROWS_5_TO_6:
            cat_val = ws_temp.cell(r, 1).value
            if cat_val: current_cat = str(cat_val).strip()
            
            desc = str(ws_temp.cell(r, 2).value or "").strip()
            weight = str(ws_temp.cell(r, WEIGHT_COL).value or "N/A").strip()
            
            item_info[r] = {
                "label": f"{current_cat} | {desc}" if current_cat not in desc else desc,
                "weight": weight
            }
        
        tab1, tab2 = st.tabs(["✍️ 組長填寫評分", "📊 彙整結算匯出"])
        
        with tab1:
            c1, c2 = st.columns(2)
            with c1: leader = st.text_input("👤 組長姓名：", placeholder="請輸入評分者名字")
            with c2: target_emp = st.selectbox("👥 評分對象：", ["-- 請選擇人員 --"] + list(employees.keys()))
            
            if leader and target_emp != "-- 請選擇人員 --":
                with st.form(f"form_{leader}_{target_emp}"):
                    st.markdown(f"### 🎯 為 **{target_emp}** 評分")
                    curr_data = st.session_state.scores.get(leader, {}).get(target_emp, {})
                    new_scores = {}
                    
                    st.subheader("一、 評量指標 (1~4項)")
                    for r in ROWS_1_TO_4:
                        st.write(f"**{item_info[r]['label']}**")
                        st.markdown(f"<span style='color:blue'>📈 分數佔比：{item_info[r]['weight']}</span>", unsafe_allow_html=True)
                        new_scores[r] = st.number_input("輸入分數", value=float(curr_data.get(r, 0.0)), format="%.3f", key=f"v_{r}")
                        st.divider()
                        
                    st.subheader("二、 加減分項 (5~6項)")
                    for r in ROWS_5_TO_6:
                        st.write(f"**{item_info[r]['label']}**")
                        st.markdown(f"<span style='color:red'>📌 加減參考：{item_info[r]['weight']}</span>", unsafe_allow_html=True)
                        new_scores[r] = st.number_input("輸入分數", value=float(curr_data.get(r, 0.0)), format="%.3f", key=f"v_{r}")
                        st.divider()
                        
                    st.subheader("三、 加扣分事蹟")
                    new_scores[ROW_TEXT] = st.text_area("說明內容", value=curr_data.get(ROW_TEXT, ""), height=150)
                    
                    if st.form_submit_button("💾 儲存評分資料", use_container_width=True):
                        if leader not in st.session_state.scores: st.session_state.scores[leader] = {}
                        st.session_state.scores[leader][target_emp] = new_scores
                        st.success(f"已儲存 {leader} 對 {target_emp} 的評分！")

        with tab2:
            if st.session_state.scores:
                st.write("📋 **已暫存評分紀錄：**")
                for l, d in st.session_state.scores.items(): st.write(f"- {l}：已評完 ({', '.join(d.keys())})")
                
                if st.button("🚀 執行 KPI 總結算並產生 Excel", type="primary", use_container_width=True):
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
                st.info("尚無評分紀錄。")
    else:
        st.info("👈 請先上傳 Excel 範本。")
