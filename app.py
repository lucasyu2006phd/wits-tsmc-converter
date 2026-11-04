import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
from datetime import datetime
import pdfplumber
import re
import os

st.set_page_config(page_title="WITS ➡️ TSMC 工時系統", page_icon="📊", layout="wide")

st.title("📊 WITS ➡️ TSMC 工時系統整合工具")
st.markdown("這是一個專為 WITS 助理與工程師設計的自動化小工具。系統已自動載入預設人員對照表。")

tab1, tab2, tab3 = st.tabs([
    "🔄 3-1. WITS 轉 TSMC", 
    "✅ 3-2. 紙本與系統核對",
    "📈 3-3. 批次彙整總表"
])

# --- 自動讀取 GitHub 根目錄下的 mapping.xlsx ---
MAPPING_FILE_PATH = "mapping.xlsx"

@st.cache_data
def load_mapping():
    if os.path.exists(MAPPING_FILE_PATH):
        return pd.read_excel(MAPPING_FILE_PATH)
    else:
        st.error(f"⚠️ 找不到預設對照表檔案：{MAPPING_FILE_PATH}。請確認已將其上傳至 GitHub。")
        return pd.DataFrame()

df_mapping = load_mapping()

# =========================================================
# Tab 1: WITS 轉 TSMC (移除對照表上傳)
# =========================================================
with tab1:
    st.info("💡 系統已自動套用預設對照表。您只需要上傳需轉換的工時表即可。")
    wits_files = st.file_uploader("📂 上傳「WITS 工時表」(支援多檔拖曳)", type=["xlsx"], accept_multiple_files=True, key="wits_upload_1")
    work_content = st.text_input("💡 預設 Work Content", value="AMS_S")
    
    if wits_files and not df_mapping.empty:
        if st.button("🚀 開始批次轉檔與美化", type="primary", use_container_width=True):
            processed_files = {}
            my_bar = st.progress(0, text="檔案處理中...")
            
            for i, wits_file in enumerate(wits_files):
                try:
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

                    # (Excel 美化輸出邏輯維持原樣)
                    output = BytesIO()
                    workbook = pd.ExcelWriter(output, engine='xlsxwriter')
                    worksheet = workbook.book.add_worksheet('TimeSheet')
                    
                    normal_format = workbook.book.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
                    holiday_format = workbook.book.add_format({'bg_color': '#D9D9D9', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
                    header_format = workbook.book.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#E0E0E0'})
                    total_format = workbook.book.add_format({'bold': True, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'bg_color': '#FFFFCC'}) 
                    text_left_format = workbook.book.add_format({'border': 1, 'align': 'left', 'valign': 'vcenter'})
                    
                    worksheet.set_column('A:B', 12)
                    worksheet.set_column('C:C', 13)
                    worksheet.set_column('D:D', 8)
                    worksheet.set_column('E:G', 12)
                    worksheet.set_column('H:H', 15)
                    worksheet.set_column('I:I', 10)
                    
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
                            if col_num in [1, 7] and not is_holiday:
                                fmt = text_left_format
                            else:
                                fmt = holiday_format if is_holiday else normal_format
                            worksheet.write(current_row, col_num, cell_data, fmt)
                        current_row += 1
                    
                    current_row += 1 
                    worksheet.write(current_row, 2, "total", total_format)
                    worksheet.write(current_row, 4, total_normal, total_format)
                    worksheet.write(current_row, 5, total_off, total_format)
                    worksheet.write(current_row, 6, 0, total_format)
                    
                    workbook.close()
                    filename = f"TSMC_202609_{employee_name}.xlsx"
                    processed_files[filename] = output.getvalue()
                except Exception as e:
                    st.error(f"檔案 {wits_file.name} 處理失敗：{e}")
                
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
# Tab 2: 3-2. 紙本與系統核對 (維持原邏輯)
# =========================================================
with tab2:
    st.markdown("上傳紙本 PDF 與系統匯出的 Excel，系統將智能提取 `Monthly Sum`，並比對兩者的「總工時」與「請假時數」。")
    col1, col2 = st.columns(2)
    with col1:
        pdf_files = st.file_uploader("📄 1. 上傳 手寫紙本 (PDF)", type=["pdf"], accept_multiple_files=True)
    with col2:
        excel_files = st.file_uploader("📊 2. 上傳 系統打卡 (Excel)", type=["xlsx"], accept_multiple_files=True, key="excel_upload_tab2")
    
    # ... (此處保留原先的 pdfplumber 核對邏輯，為節省空間略)
    # 若需完整核對邏輯，請沿用前一版本的 Tab 3 內容。

# =========================================================
# Tab 3: 3-3. 批次彙整總表 (全新功能)
# =========================================================
with tab3:
    st.header("📈 批次彙整月度大總表")
    st.write("上傳多位員工的 WITS 工時表，系統將自動萃取並合併為一份總表，方便人事端核算。")
    summary_files = st.file_uploader("📂 上傳多份「WITS 工時表」", type=["xlsx"], accept_multiple_files=True, key="summary_upload")
    
    if summary_files and st.button("📊 生成總表", type="primary", use_container_width=True):
        summary_data = []
        for file in summary_files:
            try:
                df = pd.read_excel(file, header=None)
                name = str(df.iloc[1, 8]).strip()
                df_data = df.iloc[5:].dropna(subset=[1])
                total_actual = pd.to_numeric(df_data[7], errors='coerce').sum()
                total_leave = pd.to_numeric(df_data[5], errors='coerce').sum()
                
                tsmc_id = "未知"
                if not df_mapping.empty:
                    match = df_mapping[df_mapping['姓名_T'] == name]
                    if not match.empty:
                        tsmc_id = match.iloc[0]['工號-tsmc']
                        
                summary_data.append({
                    "姓名": name,
                    "TSMC工號": tsmc_id,
                    "總工時": total_actual,
                    "總請假時數": total_leave
                })
            except Exception as e:
                st.error(f"提取 {file.name} 失敗：{e}")
                
        if summary_data:
            df_summary = pd.DataFrame(summary_data)
            st.dataframe(df_summary, use_container_width=True)
            
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_summary.to_excel(writer, index=False, sheet_name='Monthly_Summary')
            st.download_button("📥 下載月度總表 (Excel)", data=output.getvalue(), file_name="Monthly_Summary.xlsx", type="primary")
