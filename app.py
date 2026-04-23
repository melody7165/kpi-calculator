import streamlit as st
import openpyxl
from io import BytesIO

# 網頁配置
st.set_page_config(page_title="KPI 績效評分網頁系統", layout="wide")

# ==========================================
# 定義 Excel 座標 (根據您的附件精確設定)
# ==========================================
NAME_ROW = 2            # 員工姓名在第 2 列
WEIGHT_COL = 3          # 分數佔比在第 3 欄 (C)
START_EMP_COL = 4       # 第一個員工 (熊婷華) 在第 4 欄 (D)
# 指標列號定義
ROWS_1_TO_4 = [3, 4, 5, 6, 7, 8, 9, 10, 11]  # 評量指標 1~4
ROWS_5_TO_6 = [13, 14, 15, 16, 17, 18, 19]   # 加分項與扣分項 5~6
ROW_TEXT = 21                                # 加扣分事蹟文字列

def safe_float(val):
    try:
        if val is None or str(val).strip() == "": return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0

# 初始化 Session State (儲存組長評分)
if "all_scores" not in st.session_state:
    st.session_state.all_scores = {} # 結構: { "組長名": { "員工名": { 列號: 分數 } } }

# 側邊欄：模式切換
st.sidebar.title("📊 系統模式")
mode = st.sidebar.radio("切換功能：", ["模式 A：上傳 Excel 合併", "模式 B：網頁直接評分"])

# ---------------------------------------------------------
# 模式 A：上傳兩份 Excel 合併 (保留原本需求)
# ---------------------------------------------------------
if mode == "模式 A：上傳 Excel 合併":
    st.title("📂 模式 A：自動合併兩位組長的 Excel")
    col1, col2 = st.columns(2)
    with col1: f1 = st.file_uploader("上傳【組長 A】檔案", type=["xlsx"], key="a1")
    with col2: f2 = st.file_uploader("上傳【組長 B】檔案", type=["xlsx"], key="a2")

    if f1 and f2:
        if st.button("執行合併計算"):
            wb1 = openpyxl.load_workbook(f1) # 模板
            wb2 = openpyxl.load_workbook(f2, data_only=True)
            ws1, ws2 = wb1.active, wb2.active
            
            # 找到所有員工欄位 (從 D 欄開始)
            emp_cols = [c for c in range(START_EMP_COL, ws1.max_column + 1) if ws1.cell(NAME_ROW, c).value]
            
            for c in emp_cols:
                # 判斷是否兩位組長都有評分 (1-6項不全為0)
                score_sum_1 = sum(abs(safe_float(ws1.cell(r, c).value)) for r in ROWS_1_TO_4 + ROWS_5_TO_6)
                score_sum_2 = sum(abs(safe_float(ws2.cell(r, c).value)) for r in ROWS_1_TO_4 + ROWS_5_TO_6)
                
                for r in ROWS_1_TO_4:
                    v1, v2 = safe_float(ws1.cell(r, c).value), safe_float(ws2.cell(r, c).value)
                    if score_sum_1 > 0 and score_sum_2 > 0: ws1.cell(r, c).value = (v1 + v2) / 2
                    elif score_sum_2 > 0: ws1.cell(r, c).value = v2 # 若 A 沒評，取 B
                
                for r in ROWS_5_TO_6:
                    ws1.cell(r, c).value = safe_float(ws1.cell(r, c).value) + safe_float(ws2.cell(r, c).value)
                
                # 文字合併
                t1 = str(ws1.cell(ROW_TEXT, c).value or "").strip()
                t2 = str(ws2.cell(ROW_TEXT, c).value or "").strip()
                if t1 and t2: ws1.cell(ROW_TEXT, c).value = f"【組長A】:\n{t1}\n\n【組長B】:\n{t2}"
                elif t2: ws1.cell(ROW_TEXT, c).value = t2

            out = BytesIO(); wb1.save(out); out.seek(0)
            st.success("合併完成！")
            st.download_button("下載結果", out, "KPI_合併結果.xlsx")

# ---------------------------------------------------------
# 模式 B：線上直接填表評分 (全新需求)
# ---------------------------------------------------------
else:
    st.title("📝 模式 B：組長線上評分系統")
    st.sidebar.subheader("第一步：設定範本")
    template = st.sidebar.file_uploader("上傳空白 Excel 範本", type=["xlsx"])

    if template:
        wb_temp = openpyxl.load_workbook(template, data_only=True)
        ws_temp = wb_temp.active
        
        # 1. 抓取員工清單 (精確從 Column D 開始)
        employees = {}
        for c in range(START_EMP_COL, ws_temp.max_column + 1):
            name = ws_temp.cell(NAME_ROW, c).value
            if name: employees[str(name).strip()] = c
        
        # 2. 抓取項目、佔比
        item_info = {}
        for r in ROWS_1_TO_4 + ROWS_5_TO_6:
            cat = str(ws_temp.cell(r, 1).value or "").strip()
            desc = str(ws_temp.cell(r, 2).value or "").strip()
            weight = str(ws_temp.cell(r, WEIGHT_COL).value or "0").strip()
            item_info[r] = {"label": f"{cat} {desc}".strip(), "weight": weight}
            
        st.sidebar.success(f"已辨識：{len(employees)} 位員工")
        st.sidebar.write("員工名單：", ", ".join(employees.keys()))

        tab1, tab2 = st.tabs(["✍️ 組長評分填寫", "📥 結算與匯出"])
        
        with tab1:
            l_name = st.text_input("請輸入組長姓名：")
            e_name = st.selectbox("請選擇被評分員工：", ["--請選擇--"] + list(employees.keys()))
            
            if l_name and e_name != "--請選擇--":
                st.markdown(f"### 正在為 **{e_name}** 評分")
                with st.form(f"form_{l_name}_{e_name}"):
                    curr_scores = st.session_state.all_scores.get(l_name, {}).get(e_name, {})
                    new_data = {}
                    
                    st.subheader("1-4 項評量指標 (將採平均計算)")
                    for r in ROWS_1_TO_4:
                        st.write(f"📌 {item_info[r]['label']}")
                        st.caption(f"分數佔比：{item_info[r]['weight']}")
                        new_data[r] = st.number_input("評分", value=float(curr_scores.get(r, 0.0)), format="%.3f", key=f"r_{r}")
                    
                    st.subheader("5-6 項加減分 (將採累計計算)")
                    for r in ROWS_5_TO_6:
                        st.write(f"📌 {item_info[r]['label']}")
                        st.caption(f"加減分標準：{item_info[r]['weight']}")
                        new_data[r] = st.number_input("分數", value=float(curr_scores.get(r, 0.0)), format="%.3f", key=f"r_{r}")
                        
                    st.subheader("加扣分事蹟說明")
                    new_data[ROW_TEXT] = st.text_area("請詳細填寫", value=curr_scores.get(ROW_TEXT, ""), height=100)
                    
                    if st.form_submit_button("儲存這份評分"):
                        if l_name not in st.session_state.all_scores: st.session_state.all_scores[l_name] = {}
                        st.session_state.all_scores[l_name][e_name] = new_data
                        st.success(f"已儲存 {l_name} 對 {e_name} 的評分！")

        with tab2:
            if not st.session_state.all_scores:
                st.warning("目前尚無評分紀錄。")
            else:
                st.write("已收集到的組長評分：")
                for l in st.session_state.all_scores.keys():
                    st.write(f"- {l} 組長 (已評人數: {len(st.session_state.all_scores[l])})")
                
                if st.button("開始計算 KPI 並產生 Excel"):
                    wb_out = openpyxl.load_workbook(template)
                    ws_out = wb_out.active
                    
                    for emp_name, col_idx in employees.items():
                        # 找出有評分該員工的組長
                        active_leaders = []
                        for l in st.session_state.all_scores:
                            if emp_name in st.session_state.all_scores[l]:
                                # 判斷是否為無效評分 (1-6項皆為0)
                                d = st.session_state.all_scores[l][emp_name]
                                if sum(abs(d.get(r, 0)) for r in ROWS_1_TO_4 + ROWS_5_TO_6) > 0:
                                    active_leaders.append(l)
                        
                        num = len(active_leaders)
                        if num > 0:
                            # 1~4項 平均
                            for r in ROWS_1_TO_4:
                                total = sum(st.session_state.all_scores[l][emp_name].get(r, 0) for l in active_leaders)
                                ws_out.cell(r, col_idx).value = total / num
                            # 5~6項 累計
                            for r in ROWS_5_TO_6:
                                total = sum(st.session_state.all_scores[l][emp_name].get(r, 0) for l in active_leaders)
                                ws_out.cell(r, col_idx).value = total
                            # 文字合併
                            txts = [f"【{l}】:\n{st.session_state.all_scores[l][emp_name].get(ROW_TEXT, '')}" for l in active_leaders]
                            ws_out.cell(ROW_TEXT, col_idx).value = "\n\n".join(txts)
                    
                    final_out = BytesIO(); wb_out.save(final_out); final_out.seek(0)
                    st.success("計算結算完畢！")
                    st.download_button("📥 下載最終 KPI 報表", final_out, "KPI_最終網頁彙整結果.xlsx")
    else:
        st.info("👈 請先上傳 Excel 範本以顯示評分表單。")
