import streamlit as st
import openpyxl
from io import BytesIO

# 設定網頁標題與排版
st.set_page_config(page_title="部門 KPI 分數彙整系統", layout="centered")

st.title("📊 部門人員 KPI 分數彙整系統")
st.write("請上傳兩位組長的 KPI 評分 Excel 檔案 (.xlsx)，系統將自動彙整並產生包含原有樣式與算式的最終結果檔。")

# 建立檔案上傳區塊
file1 = st.file_uploader("📂 上傳【組長 A】的評分檔", type=["xlsx"])
file2 = st.file_uploader("📂 上傳【組長 B】的評分檔", type=["xlsx"])

def safe_float(val):
    """安全地將 Excel 儲存格內容轉為浮點數，若為空值或文字則視為 0"""
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def has_rated(ws, col, rows):
    """判斷該組長是否有對此員工評分 (只要 1~6 項中有任何一項不為 0 即視為有評分)"""
    for r in rows:
        if safe_float(ws.cell(row=r, column=col).value) != 0:
            return True
    return False

if file1 and file2:
    if st.button("🚀 開始彙整計算", use_container_width=True):
        try:
            # 讀取 Excel 檔案
            # wb1 作為輸出基底，不加 data_only 可「保留原本的樣式、格式與 Excel 公式」
            wb1 = openpyxl.load_workbook(file1)
            # wb2 只需要讀取純數值來進行計算
            wb2 = openpyxl.load_workbook(file2, data_only=True)
            
            ws1 = wb1.active
            ws2 = wb2.active
            
            # 找出有員工名字的欄位 (根據附件格式，第2列為員工姓名，從第 E 欄(5)開始)
            employee_cols = []
            for col in range(5, ws1.max_column + 1):
                if ws1.cell(row=2, column=col).value:
                    employee_cols.append(col)
                    
            # 定義各指標在 Excel 中的真實列號 (1-based，根據您提供的 CSV 結構推算)
            metric_1_to_4_rows = [3, 4, 5, 6, 7, 8, 9, 10, 11]  # 評量指標 1~4
            metric_5_to_6_rows = [13, 14, 15, 16, 17, 18, 19]   # 加分項與扣分項 (5~6)
            text_row = 21                                       # 加扣分事蹟文字
            
            for col in employee_cols:
                # 1. 判斷兩位組長是否有評分
                all_metric_rows = metric_1_to_4_rows + metric_5_to_6_rows
                l1_rated = has_rated(ws1, col, all_metric_rows)
                l2_rated = has_rated(ws2, col, all_metric_rows)
                
                # 2. 處理評量指標 1~4 (需要平均)
                for r in metric_1_to_4_rows:
                    v1 = safe_float(ws1.cell(row=r, column=col).value)
                    v2 = safe_float(ws2.cell(row=r, column=col).value)
                    
                    if l1_rated and l2_rated:
                        ws1.cell(row=r, column=col).value = (v1 + v2) / 2
                    elif l2_rated:  # 只有 B 有評分
                        ws1.cell(row=r, column=col).value = v2
                    # 如果只有 A 評分，保留原 v1 即可；若都沒評分，維持 0
                        
                # 3. 處理加分與扣分項 5~6 (需要累計，不平均)
                for r in metric_5_to_6_rows:
                    v1 = safe_float(ws1.cell(row=r, column=col).value)
                    v2 = safe_float(ws2.cell(row=r, column=col).value)
                    
                    if l1_rated and l2_rated:
                        ws1.cell(row=r, column=col).value = v1 + v2
                    elif l2_rated:  # 只有 B 有評分
                        ws1.cell(row=r, column=col).value = v2
                        
                # 4. 處理加扣分事蹟 (文字合併)
                t1 = ws1.cell(row=text_row, column=col).value
                t1 = str(t1).strip() if t1 else ""
                
                t2 = ws2.cell(row=text_row, column=col).value
                t2 = str(t2).strip() if t2 else ""
                
                final_text = ""
                if t1 and t2:
                    final_text = f"【組長A】:\n{t1}\n\n【組長B】:\n{t2}"
                elif t1:
                    final_text = t1
                elif t2:
                    final_text = t2
                    
                ws1.cell(row=text_row, column=col).value = final_text
                
            # 將結果儲存到記憶體中，供網頁端下載
            output = BytesIO()
            wb1.save(output)
            output.seek(0)
            
            st.success("✅ 彙整計算完成！原始格式與公式皆已保留，請點擊下方按鈕下載。")
            st.download_button(
                label="📥 下載彙整後 Excel 檔",
                data=output,
                file_name="KPI_最終彙整結果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        except Exception as e:
            st.error(f"處理檔案時發生錯誤，請確認檔案格式是否正確。錯誤訊息: {e}")