import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
from datetime import datetime
import pdfplumber
import re

st.set_page_config(page_title="緯創 ➡️ TSMC 工時系統", layout="wide")
st.title("📊 緯創 ➡️ TSMC 工時系統整合工具")

# 建立側邊欄選單
menu = st.sidebar.radio(
    "主要功能選單",
    ["🔄 1. 工時轉檔與美化 (BMS -> TSMC)", 
     "📝 2. 人員對照表管理", 
     "✅ 3. 紙本與系統工時核對"]
)

# =========================================================
# 功能一：工時表轉檔、美化 (包含假日上色、空行、總計列)
# =========================================================
if menu == "🔄 1. 工時轉檔與美化 (BMS -> TSMC)":
    st.header("🔄 工時轉檔 (支援批次與美化)")
    
    col1, col2 = st.columns(2)
    with col1:
        mapping_file = st.file_uploader("上傳「人員對照表」", type=["xlsx"])
    with col2:
        wits_files = st.file_uploader("上傳「WITS (BMS) 工時表」(可選多個)", type=["xlsx"], accept_multiple_files=True)
    
    work_content = st.text_input("預設 Work Content", value="AMS_S")
    
    if mapping_file and wits_files:
        if st.button("🚀 開始批次轉檔與美化", type="primary"):
            df_mapping = pd.read_excel(mapping_file)
            processed_files = {}
            
            for wits_file in wits_files:
                df_wits_raw = pd.read_excel(wits_file, header=None)
                employee_name = str(df_wits_raw.iloc[1, 8]).strip()
                
                # 找工號
                match = df_mapping[df_mapping['姓名_T'] == employee_name]
                if match.empty:
                    st.warning(f"找不到 {employee_name} 的 TSMC 工號，跳過此檔。")
                    continue
                tsmc_id = match.iloc[0]['工號-tsmc']
                
                # 處理工時
                df_wits_data = df_wits_raw.iloc[5:].dropna(subset=[1]).copy()
                tsmc_data = []
                total_normal = 0
                total_off = 0
                
                for _, row in df_wits_data.iterrows():
                    date_val, day_val, actual_hrs = row[1], row[2], pd.to_numeric(row[7], errors='coerce')
                    leave_hrs = pd.to_numeric(row[5], errors='coerce')
                    remark = str(row[8]) if pd.notna(row[8]) else ""
                    
                    if pd.notna(date_val):
                        date_str = date_val.strftime('%Y/%m/%d') if isinstance(date_val, datetime) else str(date_val).replace('-', '/')
                        # 判斷是否為假日
                        is_holiday = "休假日" in remark or "節" in remark or day_val in ['六', '日', 'Sat', 'Sun']
                        
                        norm_hrs = min(actual_hrs, 8) if pd.notna(actual_hrs) and actual_hrs > 0 else 0
                        off_hrs = leave_hrs if pd.notna(leave_hrs) and leave_hrs > 0 else 0
                        
                        status = "休假日" if day_val in ['六', '日', 'Sat', 'Sun'] else ("國定假日" if is_holiday and norm_hrs==0 else "")
                        work_cont = work_content if norm_hrs > 0 else ""
                        
                        tsmc_data.append([tsmc_id, employee_name, date_str, day_val, norm_hrs, off_hrs, 0, work_cont, status, "", "", "", ""])
                        total_normal += norm_hrs
                        total_off += off_hrs

                # 開始寫入 Excel (使用 xlsxwriter 做美化)
                output = BytesIO()
                workbook = pd.ExcelWriter(output, engine='xlsxwriter')
                worksheet = workbook.book.add_worksheet('TimeSheet')
                
                # 設定格式
                holiday_format = workbook.book.add_format({'bg_color': '#D9D9D9'}) # 灰色底
                normal_format = workbook.book.add_format({})
                bold_format = workbook.book.add_format({'bold': True})
                
                # 寫入前三列表頭 (符合你提供的標準格式)
                worksheet.write(0, 0, "Security C - TSMC Secret")
                worksheet.write(0, 2, "OM-20160524")
                worksheet.write(0, 3, "Please not modify the template")
                worksheet.write(1, 0, "Remark:\n This template is only for batch")
                worksheet.write(1, 2, "Format: 2016/05/01")
                worksheet.write(1, 7, "English Only \n Maximun is 2000 characters")
                
                headers = ["Employee ID", "Employee name", "Date", "Day", "Normal hours", "Off hours", "Overtime hours", "Work Content", "Status", "Approved hours", "Remark", "Submit date", "Response date"]
                for col_num, header in enumerate(headers):
                    worksheet.write(2, col_num, header, bold_format)
                
                # 寫入資料列並上色
                current_row = 3
                for row_data in tsmc_data:
                    is_holiday = row_data[8] in ["休假日", "國定假日"]
                    fmt = holiday_format if is_holiday else normal_format
                    for col_num, cell_data in enumerate(row_data):
                        worksheet.write(current_row, col_num, cell_data, fmt)
                    current_row += 1
                
                # 需求 2-4: 空留一行，再寫入 Total 列
                current_row += 1 # 刻意留空一行
                worksheet.write(current_row, 2, "total", bold_format)
                worksheet.write(current_row, 4, total_normal, bold_format)
                worksheet.write(current_row, 5, total_off, bold_format)
                worksheet.write(current_row, 6, 0, bold_format)
                
                workbook.close()
                filename = f"TSMC_202609_{employee_name}.xlsx"
                processed_files[filename] = output.getvalue()
                
            # 下載邏輯 (單檔直接下，多檔包成 ZIP)
            if processed_files:
                st.success(f"✅ 成功處理 {len(processed_files)} 份檔案並完成美化排版！")
                if len(processed_files) == 1:
                    name, data = list(processed_files.items())[0]
                    st.download_button("📥 下載轉換檔案", data=data, file_name=name)
                else:
                    zip_buf = BytesIO()
                    with zipfile.ZipFile(zip_buf, "w") as z:
                        for name, data in processed_files.items():
                            z.writestr(name, data)
                    st.download_button("📦 下載全部 (ZIP)", data=zip_buf.getvalue(), file_name="TSMC_Timesheets.zip")

# =========================================================
# 功能二：對照表管理 (可視化編輯、增加人名)
# =========================================================
elif menu == "📝 2. 人員對照表管理":
    st.header("📝 維護 TSMC 工號對照表")
    st.info("💡 你可以在下方表格直接點擊修改、新增列，完成後點擊下載，即為最新的對照表。")
    
    mapping_file = st.file_uploader("載入現有的「人員對照表」", type=["xlsx"])
    
    if mapping_file:
        df_mapping = pd.read_excel(mapping_file)
        # st.data_editor 允許使用者在網頁上直接編輯 Excel 表格內容
        edited_df = st.data_editor(df_mapping, num_rows="dynamic", use_container_width=True)
        
        output = BytesIO()
        edited_df.to_excel(output, index=False, engine='xlsxwriter')
        st.download_button("💾 下載更新後的對照表", data=output.getvalue(), file_name="wits-tsmc-工時轉換-人名工作證-對照表_新版.xlsx")

# =========================================================
# 功能三：工時核對 (PDF 手寫 vs Excel 打卡)
# =========================================================
elif menu == "✅ 3. 紙本與系統工時核對":
    st.header("✅ 紙本 (PDF) 與系統 (Excel) 工時核對")
    st.write("上傳紙本掃描的 PDF 與系統匯出的 Excel，系統會自動比對兩者的「總時數」與「請假時數」是否吻合。")
    
    pdf_files = st.file_uploader("上傳 紙本 PDF (支援批次)", type=["pdf"], accept_multiple_files=True)
    excel_files = st.file_uploader("上傳 系統 Excel (支援批次)", type=["xlsx"], accept_multiple_files=True)
    
    if st.button("🔍 開始智能核對") and pdf_files and excel_files:
        
        # 建立 Excel 總計時數的字典 (姓名: {總工時, 總請假})
        excel_summary = {}
        for ex_file in excel_files:
            df = pd.read_excel(ex_file, header=None)
            name = str(df.iloc[1, 8]).strip()
            df_data = df.iloc[5:].dropna(subset=[1])
            actual = pd.to_numeric(df_data[7], errors='coerce').sum()
            leave = pd.to_numeric(df_data[5], errors='coerce').sum()
            excel_summary[name] = {"actual": actual, "leave": leave, "filename": ex_file.name}
            
        # 解析 PDF 並比對
        for pdf_file in pdf_files:
            st.markdown("---")
            st.subheader(f"📄 核對文件：{pdf_file.name}")
            try:
                # 運用 pdfplumber 抓取 PDF 內的文字
                pdf_text = ""
                with pdfplumber.open(pdf_file) as pdf:
                    for page in pdf.pages:
                        pdf_text += page.extract_text()
                
                # 找尋 PDF 中的姓名與總計數字 (正則表達式)
                name_match = re.search(r"姓名:(\w+)", pdf_text)
                pdf_name = name_match.group(1).strip() if name_match else "未知"
                
                # 假設 PDF 底部會有 Monthly Sum: 152 8 的字樣
                sum_match = re.search(r"Monthly\s*Sum:\s*(\d+)\s*(\d+)?", pdf_text)
                pdf_actual = int(sum_match.group(1)) if sum_match else 0
                pdf_leave = int(sum_match.group(2)) if sum_match and sum_match.group(2) else 0
                
                # 比對結果
                if pdf_name in excel_summary:
                    ex_data = excel_summary[pdf_name]
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.write("**紙本 (PDF) 紀錄**")
                        st.write(f"姓名: {pdf_name}")
                        st.write(f"總工時: `{pdf_actual}` | 請假: `{pdf_leave}`")
                        
                    with col2:
                        st.write(f"**系統 (Excel) 紀錄** - `{ex_data['filename']}`")
                        st.write(f"姓名: {pdf_name}")
                        st.write(f"總工時: `{ex_data['actual']}` | 請假: `{ex_data['leave']}`")
                    
                    # 邏輯判斷
                    if pdf_actual == ex_data['actual'] and pdf_leave == ex_data['leave']:
                        st.success(f"🎉 核對成功！{pdf_name} 的紙本與系統紀錄完全一致。")
                    else:
                        st.error(f"⚠️ 異常警告！{pdf_name} 的時數有落差。")
                        diff_actual = ex_data['actual'] - pdf_actual
                        diff_leave = ex_data['leave'] - pdf_leave
                        st.write(f"> 差距分析：系統實際工時差 `{diff_actual}` 小時，請假時數差 `{diff_leave}` 小時。請人工覆核！")
                else:
                    st.warning(f"找不到 {pdf_name} 對應的 Excel 打卡紀錄。")
                    
            except Exception as e:
                st.error(f"解析 {pdf_file.name} 失敗：{e}")