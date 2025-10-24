# =====================================
# Streamlit App: 人事用“提成项目 & 二次项目 & 平台工”自动审核（多sheet版）
# =====================================

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import time

st.title("📊 人事用审核工具：起租提成 & 二次提成 & 平台工表自动检查")

# ========== 上传文件 ==========
uploaded_files = st.file_uploader(
    "请上传原始数据表（提成项目、二次明细、放款明细、本司sheet、产品台账、超期明细）",
    type="xlsx", accept_multiple_files=True
)

if not uploaded_files or len(uploaded_files) < 4:
    st.warning("⚠️ 请上传所有必要文件后继续")
    st.stop()
else:
    st.success("✅ 文件上传完成")

# ========== 工具函数 ==========
def find_file(files_list, keyword):
    for f in files_list:
        if keyword in f.name:
            return f
    raise FileNotFoundError(f"未找到包含关键词「{keyword}」的文件")

def normalize_colname(c):
    return str(c).strip().lower()

def find_col(df, keyword, exact=False):
    key = keyword.strip().lower()
    for col in df.columns:
        cname = normalize_colname(col)
        if (exact and cname == key) or (not exact and key in cname):
            return col
    return None

def normalize_num(val):
    if pd.isna(val):
        return None
    s = str(val).replace(",", "").strip()
    if s in ["", "-", "nan"]:
        return None
    try:
        if "%" in s:
            s = s.replace("%","")
            return float(s)/100
        return float(s)
    except:
        return s

def same_date_ymd(a,b):
    try:
        da = pd.to_datetime(a, errors='coerce')
        db = pd.to_datetime(b, errors='coerce')
        if pd.isna(da) or pd.isna(db):
            return False
        return (da.year==db.year) and (da.month==db.month) and (da.day==db.day)
    except:
        return False

def detect_header_row(file, sheet_name):
    """自动检测表头行位置"""
    preview = pd.read_excel(file, sheet_name=sheet_name, nrows=2, header=None)
    first_row = preview.iloc[0]
    total_cells = len(first_row)
    empty_like = sum(
        (pd.isna(x) or str(x).startswith("Unnamed") or str(x).strip() == "")
        for x in first_row
    )
    empty_ratio = empty_like / total_cells if total_cells > 0 else 0
    if empty_ratio >= 0.7:
        return 1  # 跳过备注行
    return 0

def get_header_row(file, sheet_name):
    """白名单优先：已知某些表固定header=1"""
    if any(k in sheet_name for k in ["起租", "二次"]):
        return 1
    return detect_header_row(file, sheet_name)

def compare_and_mark(idx, row, main_df, main_kw, ref_df, ref_kw, ref_contract_col, ws, red_fill, ignore_tol=0):
    errors = 0
    main_col = find_col(main_df, main_kw)
    ref_col = find_col(ref_df, ref_kw)
    if not main_col or not ref_col or not ref_contract_col:
        return 0

    contract_no = str(row.get(contract_col_main)).strip()
    if pd.isna(contract_no) or contract_no in ["", "nan"]:
        return 0

    ref_rows = ref_df[ref_df[ref_contract_col].astype(str).str.strip()==contract_no]
    if ref_rows.empty:
        return 0

    ref_val = ref_rows.iloc[0][ref_col]
    main_val = row.get(main_col)

    if pd.isna(main_val) and pd.isna(ref_val):
        return 0

    if "日期" in main_kw or "日期" in ref_kw:
        if not same_date_ymd(main_val, ref_val):
            errors = 1
    else:
        main_num = normalize_num(main_val)
        ref_num = normalize_num(ref_val)
        if isinstance(main_num,(int,float)) and isinstance(ref_num,(int,float)):
            if abs(main_num - ref_num) > ignore_tol:
                errors = 1
        else:
            if str(main_num).strip() != str(ref_num).strip():
                errors = 1

    if errors:
        excel_row = idx + 3
        col_idx = list(main_df.columns).index(main_col) + 1
        ws.cell(excel_row, col_idx).fill = red_fill
    return errors

# ========== 读取参考文件 ==========
main_file = find_file(uploaded_files, "提成项目")
ec_file = find_file(uploaded_files, "二次明细")
fk_file = find_file(uploaded_files, "放款明细")
product_file = find_file(uploaded_files, "产品台账")

ec_df = pd.read_excel(ec_file)
fk_xls = pd.ExcelFile(fk_file)
fk_df = pd.read_excel(fk_xls, sheet_name=[s for s in fk_xls.sheet_names if "本司" in s][0])
product_df = pd.read_excel(product_file)

contract_col_ec = find_col(ec_df, "合同")
contract_col_fk = find_col(fk_df, "合同")
contract_col_product = find_col(product_df, "合同")

# ========== 核心审核函数 ==========
def audit_sheet(sheet_name, main_file, ec_df, fk_df, product_df):
    xls_main = pd.ExcelFile(main_file)
    header_row = get_header_row(main_file, sheet_name)
    main_df = pd.read_excel(xls_main, sheet_name=sheet_name, header=header_row)
    st.write(f"📘 正在审核：{sheet_name}（header={header_row}）")

    contract_col_main = find_col(main_df, "合同")
    if not contract_col_main:
        st.error(f"❌ {sheet_name} 中未找到“合同号”列，已跳过。")
        return None, 0

    mappings = [
        ("起租日期", ["起租日_商","起租日_商"], 0),
        ("租赁本金", ["租赁本金"], 0),
        ("收益率", ["XIRR_商_起租"], 0.005)
    ]

    wb = Workbook()
    ws = wb.active
    for c_idx, col_name in enumerate(main_df.columns, start=1):
        ws.cell(1, c_idx, col_name)

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    total_errors = 0
    n_rows = len(main_df)
    progress = st.progress(0)
    status = st.empty()

    for idx, row in main_df.iterrows():
        contract_no = str(row.get(contract_col_main)).strip()
        if pd.isna(contract_no) or contract_no in ["", "nan"]:
            continue

        for main_kw, ref_kws, tol in mappings:
            for ref_kw, ref_df, ref_contract_col in zip(
                ref_kws,
                [ec_df, product_df] if main_kw=="起租日期" else [fk_df] if main_kw=="租赁本金" else [product_df],
                [contract_col_ec, contract_col_product] if main_kw=="起租日期" else [contract_col_fk] if main_kw=="租赁本金" else [contract_col_product]
            ):
                total_errors += compare_and_mark(idx,row,main_df,main_kw,ref_df,ref_kw,ref_contract_col,ws,red_fill,tol)

        progress.progress((idx+1)/n_rows)
        if (idx+1)%10==0 or idx+1==n_rows:
            status.text(f"{sheet_name}：{idx+1}/{n_rows} 行")

    # 标黄合同号列 & 写入数据
    contract_col_idx_excel = list(main_df.columns).index(contract_col_main)+1
    for row_idx in range(len(main_df)):
        excel_row = row_idx+3
        has_red = any(ws.cell(excel_row,c).fill==red_fill for c in range(1,len(main_df.columns)+1))
        if has_red:
            ws.cell(excel_row,contract_col_idx_excel).fill = yellow_fill
        for c_idx,val in enumerate(main_df.iloc[row_idx],start=1):
            ws.cell(excel_row,c_idx,val)

    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)

    st.download_button(
        label=f"📥 下载 {sheet_name} 审核标注版",
        data=output_stream,
        file_name=f"{sheet_name}_审核标注版.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    st.success(f"✅ {sheet_name} 审核完成，共发现 {total_errors} 处错误")

    return main_df, total_errors

# ========== 执行审核 ==========
xls_main = pd.ExcelFile(main_file)
target_sheets = [s for s in xls_main.sheet_names if any(k in s for k in ["起租", "二次", "平台工"])]

if not target_sheets:
    st.warning("⚠️ 未找到包含 '起租'、'二次' 或 '平台工' 的sheet。")
else:
    for sheet_name in target_sheets:
        audit_sheet(sheet_name, main_file, ec_df, fk_df, product_df)
