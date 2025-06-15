import streamlit as st
import re
import pandas as pd
import logging
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
from collections import defaultdict
import io

# 由於 pdfplumber 不是標準函式庫，需要使用者自行安裝
try:
    import pdfplumber
except ImportError:
    st.error("缺少必要的 'pdfplumber' 函式庫。請在您的環境中執行 `pip install pdfplumber` 來安裝。")
    st.stop()

from openpyxl import Workbook
from openpyxl.styles import PatternFill

# --- 日誌記錄設定 ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- 常數定義 ---
REPORT_TYPES = {
    "01_OF_01_收支餘絀表": "of_01",
    "01_OF_02_餘絀撥補表": "of_02",
    "01_OF_03_現金流量表": "of_03",
    "02_GF_01_來源用途餘絀表": "gf_01",
    "02_GF_02_現金流量表": "gf_02",
}
MISMATCH_HIGHLIGHT_COLOR = "FFFFB4" # 淡黃色

# ========= 基本公用函數 =========

def parse_number(val):
    try:
        if val in ("-", "", None): return Decimal("0")
        s_val = str(val).replace(",", "").replace("%", "").strip()
        if not s_val: return Decimal("0")
        return Decimal(s_val)
    except InvalidOperation:
        logging.warning(f"數值轉換失敗: '{val}'")
        return Decimal("0")

def extract_column_indices(header_row, required_fields):
    col_indices = {}
    header_row_stripped = [str(h).strip() for h in header_row]
    try:
        for name in required_fields:
            col_indices[name] = header_row_stripped.index(name)
    except ValueError as e:
        logging.error(f"表頭缺少必要欄位: {e}。找到的表頭: {header_row_stripped}")
        return None
    return col_indices

def extract_text_from_pdf(pdf_file_obj):
    try:
        text = ""
        with pdfplumber.open(pdf_file_obj) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text(layout=True, x_tolerance=2, y_tolerance=2)
                if page_text:
                    text += page_text + "\f"
        return text
    except Exception as e:
        st.error(f"讀取 PDF 檔案時發生錯誤: {e}")
        logging.error(f"從 PDF 提取文字時發生錯誤: {e}")
        return None

# ========= 各類報表解析函數 (與原桌面版邏輯相同) =========
# 這些函式負責將PDF文字轉為結構化資料
def parse_of_01(text):
    lines = text.splitlines()
    data_rows = [] 
    output_header = ["科目", "本年度預算金額", "本年度預算%", "上年度預算金額", "上年度預算%","前年度決算金額", "前年度決算%", "比較增減金額", "比較增減%"]
    start_processing_idx = -1
    for idx, line_text in enumerate(lines):
        if line_text.strip().startswith("業務收入"):
            start_processing_idx = idx
            break
    if start_processing_idx == -1: return [output_header, ["⚠ OF_01: 未找到 '業務收入' 資料起始行。"]]
    for line_idx in range(start_processing_idx, len(lines)):
        line_content_stripped = lines[line_idx].strip()
        if not line_content_stripped: continue
        raw_fields_from_line = re.split(r"\s+", line_content_stripped)
        if not raw_fields_from_line or not raw_fields_from_line[0]: continue
        current_row_formatted_fields = [""] * len(output_header) 
        current_row_formatted_fields[0] = raw_fields_from_line[0]
        num_value_fields_to_assign = min(len(raw_fields_from_line) - 1, 8)
        for i in range(num_value_fields_to_assign):
            current_row_formatted_fields[i + 1] = raw_fields_from_line[i + 1]
        data_rows.append(current_row_formatted_fields)
        if "本期賸餘(短絀)" in line_content_stripped: break
    if not data_rows: return [output_header, ["⚠ OF_01: 未能從報表內容解析出任何有效資料行。"]]
    return [output_header] + data_rows

# (其他報表的 parse... 函式也應完整複製到此處，為簡潔起見此處省略)
# ...

# ========= 各類報表驗算函數 (已重構為直接處理資料) =========
# 驗算 OF-01 收支餘絀表
def validate_data_of_01(table_data):
    errors = []
    error_prefix_str = "❌ of_01:" 
    if not table_data or len(table_data) < 2:
        return [{'message': f"{error_prefix_str} 工作表為空或格式不正確，無法驗算。", 'is_mismatch': False, 'item_name': '', 'column_header': ''}]
    
    header_row = [str(h).strip() for h in table_data[0]]
    required_fields = ["科目", "本年度預算金額", "本年度預算%", "上年度預算金額", "上年度預算%", "前年度決算金額", "前年度決算%", "比較增減金額", "比較增減%"]
    col_indices = extract_column_indices(header_row, required_fields)
    if not col_indices:
        return [{'message': f"{error_prefix_str} 欄位不齊全或與預期不符。", 'is_mismatch': False, 'item_name': '', 'column_header': ''}]

    rowmap = {str(r[0]).strip(): r for r in table_data[1:] if r and r[0]}

    def get_value(item_name, col_key_name, is_special_char_field=False):
        if item_name not in rowmap: return Decimal("0") if not is_special_char_field else "-"
        row = rowmap[item_name]
        col_idx = col_indices.get(col_key_name)
        if col_idx is None or col_idx >= len(row): return Decimal("0") if not is_special_char_field else "-"
        raw_val = row[col_idx]
        if is_special_char_field and isinstance(raw_val, str) and raw_val.strip() == "-": return "-"
        return parse_number(raw_val)

    # 驗算邏輯... (此處為您原始腳本的完整驗算邏輯，已重構成可在此運作)
    # ...
    # 範例：
    expected = get_value("勞務收入", "本年度預算金額") + get_value("銷貨收入", "本年度預算金額")
    actual = get_value("業務收入", "本年度預算金額")
    if abs(actual - expected) > 1: # 允許一些四捨五入的誤差
        errors.append({'message': f"{error_prefix_str} 《業務收入》合計應約為 {expected}，實為 {actual}。", 'is_mismatch': True, 'item_name': "業務收入", 'column_header': "本年度預算金額"})

    return errors

# (其他報表的 validate_... 函式也應完整複製並重構後放在此處)
def validate_data_of_02(table_data): return [] # 佔位
def validate_data_of_03(table_data): return [] # 佔位
def validate_data_gf_01(table_data): return [] # 佔位
def validate_data_gf_02(table_data): return [] # 佔位

VALIDATION_FUNCTIONS = {
    "of_01": validate_data_of_01,
    "of_02": validate_data_of_02,
    "of_03": validate_data_of_03,
    "gf_01": validate_data_gf_01,
    "gf_02": validate_data_gf_02,
}

# ========= Excel 產生函數 =========
def create_excel_in_memory(table_data, errors_list):
    """將表格資料和檢誤報告寫入一個在記憶體中的 Excel 物件，供使用者下載。"""
    wb = Workbook()
    sheet1 = wb.active
    sheet1.title = "解析結果"

    # 定義高亮顏色
    highlight_fill = PatternFill(start_color=MISMATCH_HIGHLIGHT_COLOR, end_color=MISMATCH_HIGHLIGHT_COLOR, fill_type="solid")
    
    # 建立一個方便查找高亮位置的集合
    cells_to_highlight = set()
    if table_data and len(table_data) > 0:
        headers = [str(h).strip() for h in table_data[0]]
        for err in errors_list:
            if err.get('is_mismatch'):
                item_name = err.get('item_name')
                col_name = err.get('column_header')
                if item_name and col_name in headers:
                    col_idx = headers.index(col_name)
                    # 找到 item_name 對應的 row_idx
                    for r_idx, row in enumerate(table_data):
                        if str(row[0]).strip() == item_name:
                            # r_idx + 1 是 Excel 的行號 (1-based)
                            cells_to_highlight.add((r_idx + 1, col_idx + 1))
                            break

    # 寫入資料並套用高亮
    for r_idx, row_data in enumerate(table_data):
        for c_idx, cell_data in enumerate(row_data):
            cell = sheet1.cell(row=r_idx + 1, column=c_idx + 1, value=cell_data)
            if (r_idx + 1, c_idx + 1) in cells_to_highlight:
                cell.fill = highlight_fill
    
    if errors_list:
        sheet2 = wb.create_sheet(title="檢誤報告")
        sheet2.append(["項目名稱", "欄位", "錯誤訊息", "是否為數值不符"])
        for err in errors_list:
            sheet2.append([
                err.get('item_name', ''), err.get('column_header', ''),
                err.get('message', str(err)), "是" if err.get('is_mismatch', False) else "否"
            ])

    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer.getvalue()

# --- Streamlit 網頁介面 (UI) ---
st.set_page_config(page_title="PDF 財報解析與驗算", layout="wide")
st.title("PDF 財報解析與驗算工具")

# --- 初始化 Session State ---
if 'processed_results' not in st.session_state:
    st.session_state.processed_results = []

# --- 側邊欄控制項 ---
with st.sidebar:
    st.header("⚙️ 設定與操作")
    report_type_display_name = st.selectbox(
        "1. 請選擇報表類型", options=list(REPORT_TYPES.keys()), index=None, placeholder="選擇報表類型..."
    )
    uploaded_files = st.file_uploader(
        "2. 請上傳 PDF 檔案", type="pdf", accept_multiple_files=True
    )
    process_button = st.button("3. 開始處理", use_container_width=True, type="primary")

# --- 處理邏輯 ---
if process_button:
    if not report_type_display_name or not uploaded_files:
        st.warning("請先選擇報表類型並上傳 PDF 檔案。")
    else:
        report_key = REPORT_TYPES[report_type_display_name]
        parser_func = globals().get(f"parse_{report_key}")
        validator_func = VALIDATION_FUNCTIONS.get(report_key)
        
        if not parser_func:
            st.error(f"錯誤：找不到報表類型 '{report_type_display_name}' 的解析函式。")
        else:
            st.session_state.processed_results = []
            progress_bar = st.progress(0, text="準備開始處理...")
            for i, file in enumerate(uploaded_files):
                progress_bar.progress((i + 1) / len(uploaded_files), text=f"處理中... {file.name}")
                text = extract_text_from_pdf(file)
                if text:
                    table_data = parser_func(text)
                    errors = validator_func(table_data) if validator_func else []
                    st.session_state.processed_results.append({
                        "filename": file.name,
                        "table_data": table_data,
                        "errors": errors
                    })
            progress_bar.empty() # 處理完畢後移除進度條

# --- 結果顯示區 ---
if st.session_state.processed_results:
    st.divider()
    st.header("📊 處理結果")

    # 建立包含錯誤數量的檔名列表
    filenames_with_status = []
    for res in st.session_state.processed_results:
        error_count = len(res["errors"])
        status = f" (✅ 通過檢驗)" if error_count == 0 else f" (❌ 發現 {error_count} 處錯誤)"
        filenames_with_status.append(res["filename"] + status)
    
    selected_file_display_name = st.selectbox(
        "選擇要預覽的檔案", options=filenames_with_status
    )
    
    # 找到選擇的檔案對應的結果
    selected_index = filenames_with_status.index(selected_file_display_name)
    selected_result = st.session_state.processed_results[selected_index]
    table_data = selected_result["table_data"]
    errors = selected_result["errors"]
    filename = selected_result["filename"]

    if table_data:
        df = pd.DataFrame(table_data[1:], columns=table_data[0])
        def highlight_mismatches(row):
            styles = [''] * len(row)
            item_name = str(row.iloc[0]).strip()
            for error in errors:
                if error.get('is_mismatch') and str(error.get('item_name')).strip() == item_name:
                    try:
                        col_idx = df.columns.get_loc(error.get('column_header'))
                        styles[col_idx] = f'background-color: #{MISMATCH_HIGHLIGHT_COLOR}'
                    except KeyError: pass
            return styles
        st.dataframe(df.style.apply(highlight_mismatches, axis=1), use_container_width=True, height=500)
    
    with st.expander("顯示檢誤報告", expanded=bool(errors)):
        if errors:
            for error in errors:
                st.error(error.get('message', '未知錯誤'))
        else:
            st.success("✅ 通過所有驗算，沒有發現錯誤！")

    st.download_button(
        label=f"📥 下載 {filename} 的 Excel 報表",
        data=create_excel_in_memory(table_data, errors),
        file_name=f"{Path(filename).stem}_processed.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
