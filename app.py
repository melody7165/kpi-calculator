import streamlit as st
import openpyxl
from io import BytesIO

# 設定網頁標題與寬版排版
st.set_page_config(page_title="KPI 績效管理系統 v4", layout="wide")

# 根據 26_Q1-3_個人績效評量_空.xlsx 定義固定列號
ROWS_1_TO_4 = [3, 4, 5, 6, 7, 8, 9, 10, 11]  # 評量指標 1~4
ROWS_5_TO_6 = [13, 14, 15, 16, 17, 18, 19]   # 加減分項 5~6
ROW_TEXT = 21                                # 加扣分事蹟說明
NAME_ROW = 2                                 # 人員姓名與標題所在列

def safe_float(val):
    try:
        if val is None or str(val).strip() == "": return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0

# 側邊欄：功能模式切換
st.sidebar.title("🛠️ KPI 系統管理控制台")
app_mode = st.sidebar.radio("請選擇操作方式：", ["模式 A：合併兩份 Excel 檔", "模式 B：線上直接填表評分"])

# ---------------------------------------------------------
# 模式 A：兩份檔案自動合併 (原始上傳功能)
# ---------------------------------------------------------
if app_mode == "模式 A：合併兩份 Excel 檔":
    st.title("📂 模式 A：兩份 Excel 自動彙整")
    st.write("請分別上傳兩位組長的評分原檔，系統將自動平均/累加分數。")
    
    col1, col2 = st.columns(2)
    with col1: f1 = st.file_uploader("📂 上傳【組長 A】檔案", type=["xlsx"], key="modeA_1")
    with col2: f2 = st.file_uploader("📂 上傳【組長 B】檔案", type=["xlsx"], key="modeA_2")

    if f1 and f2:
        if st.button("🚀 開始彙整並保留公式樣式", use_container_width=True):
            wb1 = openpyxl.load_workbook(f1) # 保留樣式模板
            wb2 = openpyxl.load_workbook(f2, data_only=True)
            ws1, ws2 = wb1.active, wb2.active
            
            # 自動偵測人員欄位：從「分數佔比」欄位的下一欄開始
            weight_col = 3 # 預設 C 欄
            for c in range(1, 10):
                if "分數佔比" in str(ws1.cell(NAME_ROW, c).value or ""):
                    weight_col = c
                    break
            
            emp_cols = [c for c in range(weight_col + 1, ws1.max_column + 1) if ws1.cell(NAME_ROW, c).value]
            
            for c in emp_cols:
                # 1~4項 平均 (若單方沒填則不平均)
                for r in ROWS_1_TO_4:
                    v1, v2 = safe_float(ws1.cell(r, c).value), safe_float(ws2.cell(r, c).value)
                    if v1 != 0 and v2 != 0: ws1.cell(r, c).value = (v1 + v2) / 2
                    elif v2 != 0: ws1.cell(r, c).value = v2
                # 5~6項 直接累加
                for r in ROWS_5_TO_6:
                    ws1.cell(r, c).value = safe_float(ws1.cell(r, c).value) + safe_float(ws2.cell(r, c).value)
                # 事蹟合併
                t1 = str(ws1.cell(ROW_TEXT, c).value or "").strip()
                t2 = str(ws2.cell(ROW_TEXT, c).value or "").strip()
                if t1 and t2: ws1.cell(ROW_TEXT, c).value = f"【組長A】:\n{t1}\n\n【組長B】:\n{t2}"
                elif t2: ws1.cell(ROW_TEXT, c).value = t2

            out = BytesIO(); wb1.save(out); out.seek(0)
            st.success("✅ 合併結算完成！")
            st.download_button("📥 下載最終 Excel 檔", out, "KPI_彙整結果.xlsx")

# ---------------------------------------------------------
# 模式 B：線上評分系統 (根據最新範本校準)
# ---------------------------------------------------------
elif app_mode == "模式 B：線上直接填表評分":
    if "scores" not in st.session_state: st.session_state.scores = {}

    st.title("📝 模式 B：線上評分與結算系統")
    st.sidebar.divider()
    temp_file = st.sidebar.file_uploader("📂 第一步：上傳空白 Excel 範本", type=["xlsx"], key="modeB_temp")

    if temp_file:
        wb_temp = openpyxl.load_workbook(temp_file, data_only=True)
        ws_temp = wb_temp.active
        
        # 1. 動態定位「分數佔比」欄位與「人員起始」欄位
        weight_col_idx = 3
        for c in range(1, 15):
            val = str(ws_temp.cell(NAME_ROW, c).value or "")
            if "分數佔比" in val:
                weight_col_idx = c
                break
        
        # 2. 抓取所有員工 (從佔比欄的下一欄開始，確保 熊婷華 在內)
        employees = {}
        for c in range(weight_col_idx + 1, ws_temp.max_column + 1):
            name = ws_temp.cell(row=NAME_ROW, column=c).value
            if name: employees[str(name).strip()] = c
        
        st.sidebar.write("✅ **系統已成功辨識人員：**")
        st.sidebar.code("、".join(employees.keys()))
        
        # 3. 抓取評分指標與正確佔比
        item_info = {}
        curr_cat = ""
        for r in ROWS_1_TO_4 + ROWS_5_TO_6:
            cat = ws_temp.cell(r, 1).value
            if cat: curr_cat = str(cat).strip()
            desc = str(ws_temp.cell(r, 2).value or "").strip()
            weight = str(ws_temp.cell(r, weight_col_idx).value or "0").strip() # 抓取對應的 C 欄數值
            
            item_info[r] = {
                "title": f"{curr_cat} | {desc}" if curr_cat not in desc else desc,
                "weight": weight
            }
        
        t1, t2 = st.tabs(["✍️ 開始評分", "📊 數據彙整匯出"])
        
        with t1:
            l_col, e_col = st.columns(2)
            with l_col: leader = st.text_input("👤 組長姓名：", placeholder="請輸入評分者名字")
            with e_col: target_e = st.selectbox("👥 評分對象：", ["-- 請選擇人員 --"] + list(employees.keys()))
            
            if leader and target_e != "-- 請選擇人員 --":
                with st.form(f"form_{leader}_{target_e}"):
                    st.write(f"### 🎯 為 **{target_e}** 評分")
                    old_data = st.session_state.scores.get(leader, {}).get(target_e, {})
                    new_scores = {}
                    
                    st.subheader("一、 評量指標 (1~4項)")
                    for r in ROWS_1_TO_4:
                        st.markdown(f"**{item_info[r]['title']}**")
                        st.markdown(f"<p style='color:#007BFF; font-size:14px;'>⚖️ 本項佔比：{item_info[r]['weight']}</p>", unsafe_allow_html=True)
                        new_scores[r] = st.number_input("輸入得分", value=float(old_data.get(r, 0.0)), format="%.3f", key=f"val_{r}")
                        st.divider()
                        
                    st.subheader("二、 加減分項目 (5~6項)")
                    for r in ROWS_5_TO_6:
                        st.markdown(f"**{item_info[r]['title']}**")
                        st.markdown(f"<p style='color:#DC3545; font-size:14px;'>📌 參考基準：{item_info[r]['weight']}</p>", unsafe_allow_html=True)
                        new_scores[r] = st.number_input("輸入得分", value=float(old_data.get(r, 0.0)), format="%.3f", key=f"val_{r}")
                        st.divider()
                        
                    st.subheader("三、 具體事蹟說明")
                    new_scores[ROW_TEXT] = st.text_area("請填寫加扣分事蹟說明", value=old_data.get(ROW_TEXT, ""), height=150)
                    
                    if st.form_submit_button("💾 儲存評分紀錄", use_container_width=True):
                        if leader not in st.session_state.scores: st.session_state.scores[leader] = {}
                        st.session_state.scores[leader][target_e] = new_scores
                        st.success(f"✅ 已成功記錄 {leader} 對 {target_e} 的評分。")

        with t2:
            if st.session_state.scores:
                st.write("📋 **目前系統內的評分進度：**")
                for l, d in st.session_state.scores.items(): st.write(f"✅ {l} 組長：已完成 ({', '.join(d.keys())})")
                
                if st.button("🚀 執行 KPI 總結算並產生 Excel 檔", type="primary", use_container_width=True):
                    wb_out = openpyxl.load_workbook(template_file)
                    ws_out = wb_out.active
                    for emp, col_idx in employees.items():
                        active_ls = [l for l in st.session_state.scores if emp in st.session_state.scores[l]]
                        num = len(active_ls)
                        if num > 0:
                            for r in ROWS_1_TO_4: ws_out.cell(r, col_idx).value = sum(st.session_state.scores[l][emp].get(r, 0.0) for l in active_ls) / num
                            for r in ROWS_5_TO_6: ws_out.cell(r, col_idx).value = sum(st.session_state.scores[l][emp].get(r, 0.0) for l in active_ls)
                            ws_out.cell(ROW_TEXT, col_idx).value = "\n\n".join([f"【{l}】:\n{st.session_state.scores[l][emp].get(ROW_TEXT, '')}" for l in active_ls])
                    
                    final_out = BytesIO(); wb_out.save(final_out); final_out.seek(0)
                    st.success("🎉 全員結算完成！")
                    st.download_button("📥 下載結算報表", final_out, "KPI_線上評分結算結果.xlsx", use_container_width=True)
            else:
                st.info("尚無任何評分數據。")
    else:
        st.info("👈 請先上傳 Excel 範本以啟用線上評分表單。")
