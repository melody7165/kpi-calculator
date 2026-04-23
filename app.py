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

def format_to_percentage_string(val):
    """將 Excel 中的小數佔比轉為百分比字串 (例如 0.3 轉為 30%)"""
    try:
        if val is None or str(val).strip() == "": return "0%"
        f_val = float(val)
        return f"{f_val * 100:g}%"
    except ValueError:
        return str(val).strip() # 如果是 +X% 這類的文字，就直接保留

# 初始化 Session State (儲存組長評分，以小數點格式儲存)
if "all_scores" not in st.session_state:
    st.session_state.all_scores = {}

# 側邊欄：模式切換
st.sidebar.title("📊 系統模式")
mode = st.sidebar.radio("切換功能：", ["模式 A：上傳 Excel 合併", "模式 B：網頁直接評分"])

# ---------------------------------------------------------
# 模式 A：上傳兩份 Excel 合併
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
            
            emp_cols = [c for c in range(START_EMP_COL, ws1.max_column + 1) if ws1.cell(NAME_ROW, c).value]
            
            for c in emp_cols:
                score_sum_1 = sum(abs(safe_float(ws1.cell(r, c).value)) for r in ROWS_1_TO_4 + ROWS_5_TO_6)
                score_sum_2 = sum(abs(safe_float(ws2.cell(r, c).value)) for r in ROWS_1_TO_4 + ROWS_5_TO_6)
                
                for r in ROWS_1_TO_4:
                    v1, v2 = safe_float(ws1.cell(r, c).value), safe_float(ws2.cell(r, c).value)
                    if score_sum_1 > 0 and score_sum_2 > 0: ws1.cell(r, c).value = (v1 + v2) / 2
                    elif score_sum_2 > 0: ws1.cell(r, c).value = v2 
                
                for r in ROWS_5_TO_6:
                    ws1.cell(r, c).value = safe_float(ws1.cell(r, c).value) + safe_float(ws2.cell(r, c).value)
                
                t1 = str(ws1.cell(ROW_TEXT, c).value or "").strip()
                t2 = str(ws2.cell(ROW_TEXT, c).value or "").strip()
                if t1 and t2: ws1.cell(ROW_TEXT, c).value = f"【組長A】:\n{t1}\n\n【組長B】:\n{t2}"
                elif t2: ws1.cell(ROW_TEXT, c).value = t2

            out = BytesIO(); wb1.save(out); out.seek(0)
            st.success("合併完成！")
            st.download_button("下載結果", out, "KPI_合併結果.xlsx")

# ---------------------------------------------------------
# 模式 B：線上直接填表評分
# ---------------------------------------------------------
else:
    st.title("📝 模式 B：組長線上評分系統")
    st.sidebar.subheader("第一步：設定範本")
    template = st.sidebar.file_uploader("上傳空白 Excel 範本", type=["xlsx"])

    if template:
        wb_temp = openpyxl.load_workbook(template, data_only=True)
        ws_temp = wb_temp.active
        
        employees = {}
        for c in range(START_EMP_COL, ws_temp.max_column + 1):
            name = ws_temp.cell(NAME_ROW, c).value
            if name: employees[str(name).strip()] = c
        
        item_info = {}
        for r in ROWS_1_TO_4 + ROWS_5_TO_6:
            cat = str(ws_temp.cell(r, 1).value or "").strip()
            desc = str(ws_temp.cell(r, 2).value or "").strip()
            weight_raw = ws_temp.cell(r, WEIGHT_COL).value
            item_info[r] = {
                "label": f"{cat} {desc}".strip(), 
                "weight_str": format_to_percentage_string(weight_raw) # 轉換為 30% 字串格式
            }
            
        st.sidebar.success(f"已辨識：{len(employees)} 位員工")
        st.sidebar.write("員工名單：", ", ".join(employees.keys()))

        tab1, tab2 = st.tabs(["✍️ 組長評分填寫", "📥 結算與匯出"])
        
        with tab1:
            l_name = st.text_input("請輸入組長姓名：")
            e_name = st.selectbox("請選擇被評分員工：", ["--請選擇--"] + list(employees.keys()))
            
            if l_name and e_name != "--請選擇--":
                st.markdown(f"### 正在為 **{e_name}** 評分")
                st.info("💡 **貼心提示：** 請直接輸入百分比數字。例如想給 8.5% 分，請直接在格子輸入 **8.5** 即可。")
                with st.form(
