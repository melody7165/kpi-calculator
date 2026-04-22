import streamlit as st
import openpyxl
from io import BytesIO

# 設定網頁標題與排版
st.set_page_config(page_title="KPI 績效管理系統", layout="wide")

# 常用設定與計算邏輯 (根據附件 Excel 結構)
ROWS_1_TO_4 = [3, 4, 5, 6, 7, 8, 9, 10, 11]  # 需要平均的項目
ROWS_5_TO_6 = [13, 14, 15, 16, 17, 18, 19]   # 需要累加的項目
ROW_TEXT = 21                                # 加扣分事蹟說明列

def safe_float(val):
    try:
        if val is None: return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0

# 側邊欄功能切換
st.sidebar.title("🛠️ 功能切換")
app_mode = st.sidebar.radio("請選擇操作方式：", ["模式 A：合併兩份 Excel 檔", "模式 B：線上直接填表評分"])

# ---------------------------------------------------------
# 模式 A：兩份檔案自動合併 (原始功能)
# ---------------------------------------------------------
if app_mode == "模式 A：合併兩份 Excel 檔":
    st.title("📂 模式 A：兩份檔案自動合併")
    st.write("適合組長已填好各自的 Excel 檔案，您需要快速合併時使用。")
    
    col1, col2 = st.columns(2)
    with col1:
        f1 = st.file_uploader("📂 上傳【組長 A】評分檔", type=["xlsx"], key="m_f1")
    with col2:
        f2 = st.file_uploader("📂 上傳【組長 B】評分檔", type=["xlsx"], key="m_f2")

    if f1 and f2:
        if st.button("🚀 開始彙整檔案", use_container_width=True):
            try:
                wb1 = openpyxl.load_workbook(f1) # 以 A 為模板
                wb2 = openpyxl.load_workbook(f2, data_only=True)
                ws1, ws2 = wb1.active, wb2.active
                
                # 【修正】新版 Excel: 員工名字從第 4 欄 (D欄) 開始
                emp_cols = [c for c in range(4, ws1.max_column + 1) if ws1.cell(2, c).value]
                
                for c in emp_cols:
                    # 1~4項 平均 (考慮單方未評分狀況)
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

                output = BytesIO()
                wb1.save(output)
                output.seek(0)
                st.success("✅ 合併完成！")
                st.download_button("📥 下載彙整結果", output, "KPI_合併結果.xlsx")
            except Exception as e:
                st.error(f"檔案處理失敗，請確認上傳的是正確的 Excel 檔。錯誤訊息：{e}")

# ---------------------------------------------------------
# 模式 B：線上評分系統
# ---------------------------------------------------------
elif app_mode == "模式 B：線上直接填表評分":
    if "scores" not in st.session_state:
        st.session_state.scores = {}

    st.title("📝 模式 B：線上直接填表評分")
    st.sidebar.divider()
    st.sidebar.subheader("第一步：設定範本")
    template_file = st.sidebar.file_uploader("📂 上傳空白 Excel 範本", type=["xlsx"], key="t_file")

    if template_file:
        try:
            wb_temp = openpyxl.load_workbook(template_file, data_only=True)
            ws_temp = wb_temp.active
            
            # 1. 抓取員工名單 【修正】新版從第 4 欄 (D欄) 開始
            employees = {str(ws_temp.cell(2, c).value).strip(): c for c in range(4, ws_temp.max_column + 1) if ws_temp.cell(2, c).value}
            
            # 2. 抓取評分項目資訊 【修正】佔比固定在第 3 欄 (C欄)
            item_info = {}
            for r in ROWS_1_TO_4 + ROWS_5_TO_6:
                cat = str(ws_temp.cell(r, 1).value or "").strip()
                desc = str(ws_temp.cell(r, 2).value or "").strip()
                weight = str(ws_temp.cell(r, 3).value or "N/A").strip() # 改為抓取第三欄 (C欄)
                item_info[r] = {
                    "label": f"{cat} {desc}",
                    "weight": weight
                }
            
            tab1, tab2 = st.tabs(["✍️ 組長評分區", "📊 結算匯出區"])
            
            with tab1:
                col_l, col_e = st.columns(2)
                with col_l: l_name = st.text_input("👤 組長姓名：", placeholder="請輸入您的名字")
                with col_e: e_name = st.selectbox("👥 評分對象：", ["-- 請選擇 --"] + list(employees.keys()))
                
                if l_name and e_name != "-- 請選擇 --":
                    with st.form(f"form_{l_name}_{e_name}"):
                        curr_saved = st.session_state.scores.get(l_name, {}).get(e_name, {})
                        new_entry = {}
                        
                        st.markdown(f"### 🎯 正在評分：{e_name}")
                        st.info("💡 項目名稱下方會顯示該項目的【分數佔比】，請依此標準給分。")
                        
                        st.subheader("一、 評量指標 (1~4項)")
                        for r in ROWS_1_TO_4:
                            st.write(f"**{item_info[r]['label']}**")
                            st.caption(f"📌 分數佔比：{item_info[r]['weight']}") # 顯示正確佔比
                            new_entry[r] = st.number_input("請輸入分數", value=float(curr_saved.get(r, 0.0)), format="%.3f", key=f"in_{r}")
                            st.divider()
                            
                        st.subheader("二、 加減分項 (5~6項)")
                        for r in ROWS_5_TO_6:
                            st.write(f"**{item_info[r]['label']}**")
                            st.caption(f"📌 加減分參考：{item_info[r]['weight']}")
                            new_entry[r] = st.number_input("請輸入分數", value=float(curr_saved.get(r, 0.0)), format="%.3f", key=f"in_{r}")
                            st.divider()
                            
                        st.subheader("三、 加扣分事蹟說明")
                        new_entry[ROW_TEXT] = st.text_area("請詳細填寫事蹟內容", value=curr_saved.get(ROW_TEXT, ""), height=150)
                        
                        if st.form_submit_button("💾 儲存此員工評分", use_container_width=True):
                            if l_name not in st.session_state.scores: st.session_state.scores[l_name] = {}
                            st.session_state.scores[l_name][e_name] = new_entry
                            st.success(f"✅ 已成功儲存 {l_name} 對 {e_name} 的評分！")

            with tab2:
                if st.session_state.scores:
                    st.write("📋 **目前已暫存的評分清單：**")
                    for l, d in st.session_state.scores.items():
                        st.write(f"- **{l}** 組長：已完成 ({', '.join(d.keys())})")
                    
                    if st.button("🚀 執行 KPI 全員結算並匯出 Excel", type="primary", use_container_width=True):
                        wb_out = openpyxl.load_workbook(template_file)
                        ws_out = wb_out.active
                        
                        for emp, col_idx in employees.items():
                            v_leaders = [l for l in st.session_state.scores if emp in st.session_state.scores[l]]
                            num = len(v_leaders)
                            
                            if num > 0:
                                # 1~4項 平均
                                for r in ROWS_1_TO_4:
                                    ws_out.cell(r, col_idx).value = sum(st.session_state.scores[l][emp].get(r, 0.0) for l in v_leaders) / num
                                # 5~6項 累加
                                for r in ROWS_5_TO_6:
                                    ws_out.cell(r, col_idx).value = sum(st.session_state.scores[l][emp].get(r, 0.0) for l in v_leaders)
                                # 文字合併
                                ws_out.cell(ROW_TEXT, col_idx).value = "\n\n".join([f"【{l}】:\n{st.session_state.scores[l][emp].get(ROW_TEXT, '')}" for l in v_leaders])
                        
                        out_b = BytesIO()
                        wb_out.save(out_b)
                        out_b.seek(0)
                        st.success("🎉 結算完成！")
                        st.download_button("📥 下載最終 KPI 報表", out_b, "KPI_線上結算最終結果.xlsx", use_container_width=True)
                else:
                    st.info("目前尚無任何組長提交評分資料。")
        except Exception as e:
            st.error(f"範本讀取失敗，請確認上傳的是正確的 Excel 檔。錯誤訊息：{e}")
    else:
        st.info("👈 請先在左側邊欄上傳 Excel 範本，系統才能讀取評分項目。")
