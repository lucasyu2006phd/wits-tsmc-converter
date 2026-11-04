import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
from datetime import datetime
import pdfplumber
import re

# 網頁基本設定 (使用寬版配置)
st.set_page_config(page_title="WITS ➡️ TSMC 工時系統", page_icon="📊", layout="wide")

st.title("📊 WITS ➡️ TSMC 工時系統整合工具")
st.markdown("這是一個專為 WITS 助理與工程師設計的自動化小工具。")

# UI 優化 1：使用 Tabs (分頁) 讓介面更現代、更乾淨
tab1, tab2, tab3 = st.tabs([
    "🔄 1. 批次轉檔與美化", 
    "📝 2. 人員對照表管理", 
    "✅ 3. 智能工時核對"
])

# =========================================================
# Tab 1: 工時轉檔與美化
# =========================================================
with tab1:
    # UI 優化 2：將使用說明收納至 Expander，維持畫面簡潔
    with st.expander("📌 點我看【工時轉檔】使用說明"):
        st.write("""
        1. 準備好**人員對照表** (包含姓名與工號對應)。
        2. 從 BMS 系統匯出該月的所有 WITS 工時表。
        3. 將檔案拖曳至下方，點擊「開始批次轉檔與美化」，系統會自動處理假日標色與 Excel 排版。
        """)

    col1, col2 = st.columns(2)
    with col1:
        mapping_file = st.file_uploader("📂 1. 上傳「人員對照表」", type=["xlsx"])
    with col2:
        wits_files = st.file_uploader("📂 2. 上傳「WITS 工時表」(支援多檔拖曳)", type=["xlsx"], accept_multiple_files=True)
    
    work_content = st.text_input("💡 預設 Work Content", value="AMS_S")
    
    if mapping_file and wits_files:
        if st.button("🚀 開始批次轉檔與美化", type="primary", use_container_width=True):
            df_mapping = pd.read_excel(mapping_file)
            processed_files = {}
            
            # UI 優化 3：進度條
            progress_text = "檔案處理中，請稍候..."
            my_bar = st.progress(0, text=progress_text)
            
            for i, wits_file in enumerate(wits_files):
                df_wits_raw = pd.read_excel(wits_file, header=None)
                employee_name = str(df_wits_raw.iloc[1, 8]).strip()
                
                match = df_mapping[df_mapping['姓名_T'] == employee_name]
                if match.empty:
                    st.toast(f"⚠️ 略過 {employee_name}：找不到 TSMC 工號", icon="⚠️")
                    continue
                tsmc_id = match.iloc[0]['工號-tsmc']
                
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
                        is_holiday = "休假日" in remark or "節" in remark or day_val in ['六', '日', 'Sat', 'Sun']
                        
                        norm_hrs = min(actual_hrs, 8) if pd.notna(actual_hrs) and actual_hrs > 0 else 0
                        off_hrs = leave_hrs if pd.notna(leave_hrs) and leave_hrs > 0 else 0
                        
                        status = "休假日" if day_val in ['六', '日', 'Sat', 'Sun'] else ("國定假日" if is_holiday and norm_hrs==0 else "")
                        work_cont = work_content if norm_hrs > 0 else ""
                        
                        tsmc_data.append([tsmc_id, employee_name, date_str, day_val, norm_hrs, off_hrs, 0, work_cont, status, "", "", "", ""])
                        total_normal += norm_hrs
                        total_off += off_hrs

                # === Excel 輸出表格美化 (XlsxWriter) ===
                output = BytesIO()
                workbook = pd.ExcelWriter(output, engine='xlsxwriter')
                worksheet = workbook.book.add_worksheet('TimeSheet')
                
                # 設定儲存格格式 (置中、框線、背景色)
                normal_format = workbook.book.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                holiday_format = workbook.book.add_format({'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                header_format = workbook.book.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0'})
                total_format = workbook.book.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFFCC'}) # 亮黃色
                text_left_format = workbook.book.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
                
                # 設定欄寬最佳化
                worksheet.set_column('A:B', 12)  # ID, Name
                worksheet.set_column('C:C', 13)  # Date
                worksheet.set_column('D:D', 8)   # Day
                worksheet.set_column('E:G', 12)  # Hours
                worksheet.set_column('H:H', 15)  # Content
                worksheet.set_column('I:I', 10)  # Status
                
                # 凍結前三列表頭
                worksheet.freeze_panes(3, 0)
                
                worksheet.write(0, 0, "Security C - TSMC Secret")
                worksheet.write(0, 2, "OM-20160524")
                worksheet.write(0, 3, "Please not modify the template")
                worksheet.write(1, 0, "Remark:\n This template is only for batch")
                worksheet.write(1, 2, "Format: 2016/05/01")
                worksheet.write(1, 7, "English Only \n Maximun is 2000 characters")
                
                headers = ["Employee ID", "Employee name", "Date", "Day", "Normal hours", "Off hours", "Overtime hours", "Work Content", "Status", "Approved hours", "Remark", "Submit date", "Response date"]
                for col_num, header in enumerate(headers):
                    worksheet.write(2, col_num, header, header_format)
                
                current_row = 3
                for row_data in tsmc_data:
                    is_holiday = row_data[8] in ["休假日", "國定假日"]
                    for col_num, cell_data in enumerate(row_data):
                        # Work Content 和 姓名靠左，其餘置中
                        if col_num in [1, 7] and not is_holiday:
                            fmt = text_left_format
                        else:
                            fmt = holiday_format if is_holiday else normal_format
                        worksheet.write(current_row, col_num, cell_data, fmt)
                    current_row += 1
                
                current_row += 1 # 留空一行
                worksheet.write(current_row, 2, "total", total_format)
                worksheet.write(current_row, 4, total_normal, total_format)
                worksheet.write(current_row, 5, total_off, total_format)
                worksheet.write(current_row, 6, 0, total_format)
                
                workbook.close()
                filename = f"TSMC_202609_{employee_name}.xlsx"
                processed_files[filename] = output.getvalue()
                
                # 更新進度條
                my_bar.progress((i + 1) / len(wits_files), text=f"處理進度: {i+1}/{len(wits_files)}")
            
            if processed_files:
                st.success(f"🎉 成功處理並美化 {len(processed_files)} 份工時表！", icon="✅")
                
                if len(processed_files) == 1:
                    name, data = list(processed_files.items())[0]
                    st.download_button("📥 下載轉換檔案", data=data, file_name=name, type="primary")
                else:
                    zip_buf = BytesIO()
                    with zipfile.ZipFile(zip_buf, "w") as z:
                        for name, data in processed_files.items():
                            z.writestr(name, data)
                    st.download_button("📦 下載全部打包 (ZIP)", data=zip_buf.getvalue(), file_name="TSMC_Timesheets.zip", type="primary")

# =========================================================
# Tab 2: 對照表管理
# =========================================================
with tab2:
    st.info("💡 **線上編輯模式**：直接點擊下方表格修改，或是往下拉新增列，完成後點擊下載即可存檔。")
    mapping_file = st.file_uploader("📂 載入現有的「人員對照表」", type=["xlsx"], key="mapping_upload")
    
    if mapping_file:
        df_mapping = pd.read_excel(mapping_file)
        # UI 優化 4：將表格放到介面中，並填滿寬度
        edited_df = st.data_editor(df_mapping, num_rows="dynamic", use_container_width=True, height=400)
        
        output = BytesIO()
        edited_df.to_excel(output, index=False, engine='xlsxwriter')
        st.download_button("💾 儲存並下載新版對照表", data=output.getvalue(), file_name="wits-tsmc-工時轉換-人名對照表.xlsx")

# =========================================================
# Tab 3: 紙本與系統工時核對
# =========================================================
with tab3:
    st.markdown("上傳紙本 PDF 與系統匯出的 Excel，系統將智能提取 `Monthly Sum`，並比對兩者的「總工時」與「請假時數」。")
    col1, col2 = st.columns(2)
    with col1:
        pdf_files = st.file_uploader("📄 1. 上傳 手寫紙本 (PDF)", type=["pdf"], accept_multiple_files=True)
    with col2:
        excel_files = st.file_uploader("📊 2. 上傳 系統打卡 (Excel)", type=["xlsx"], accept_multiple_files=True)
    
    if st.button("🔍 開始智能核對", type="primary", use_container_width=True) and pdf_files and excel_files:
        excel_summary = {}
        for ex_file in excel_files:
            df = pd.read_excel(ex_file, header=None)
            name = str(df.iloc[1, 8]).strip()
            df_data = df.iloc[5:].dropna(subset=[1])
            actual = pd.to_numeric(df_data[7], errors='coerce').sum()
            leave = pd.to_numeric(df_data[5], errors='coerce').sum()
            excel_summary[name] = {"actual": actual, "leave": leave, "filename": ex_file.name}
            
        for pdf_file in pdf_files:
            st.divider() # 使用水平線切割每個人的結果
            try:
                pdf_text = ""
                with pdfplumber.open(pdf_file) as pdf:
                    for page in pdf.pages:
                        pdf_text += page.extract_text()
                
                name_match = re.search(r"姓名:(\w+)", pdf_text)
                pdf_name = name_match.group(1).strip() if name_match else "未知"
                
                sum_match = re.search(r"Monthly\s*Sum:\s*(\d+)\s*(\d+)?", pdf_text)
                pdf_actual = int(sum_match.group(1)) if sum_match else 0
                pdf_leave = int(sum_match.group(2)) if sum_match and sum_match.group(2) else 0
                
                if pdf_name in excel_summary:
                    ex_data = excel_summary[pdf_name]
                    
                    if pdf_actual == ex_data['actual'] and pdf_leave == ex_data['leave']:
                        st.success(f"✅ **{pdf_name}**：核對成功！紙本與系統紀錄完全一致。")
                    else:
                        st.error(f"❌ **{pdf_name}**：發現異常差異！")
                        
                        # UI 優化 5：使用 Metrics 清楚展示差異數字
                        mcol1, mcol2, mcol3 = st.columns(3)
                        mcol1.metric("來源", "紙本 (PDF)", "手寫")
                        mcol2.metric("總工時比較", f"{pdf_actual} hr", delta=f"{pdf_actual - ex_data['actual']} hr (與系統差)", delta_color="inverse")
                        mcol3.metric("請假時數比較", f"{pdf_leave} hr", delta=f"{pdf_leave - ex_data['leave']} hr (與系統差)", delta_color="inverse")
                        
                        st.caption(f"對應的系統檔案：`{ex_data['filename']}`")
                else:
                    st.warning(f"找不到 **{pdf_name}** 對應的 Excel 打卡紀錄。")
                    
            except Exception as e:
                st.error(f"解析 {pdf_file.name} 失敗：{e}")
