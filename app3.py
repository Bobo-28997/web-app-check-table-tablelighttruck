# =====================================
# Streamlit App: 人事用“提成项目”起租提成审核（更新版）
# =====================================

import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import time

st.title("📊 人事用审核工具：起租提成表自动检查（更新版）")

# 上传文件
uploaded_files = st.file_uploader(
    "请上传原始数据表（提成项目、二次明细、放款明细、产品台账）",
    type="xlsx", accept_multiple_files=True
)

if not uploaded_files or len(uploaded_files) < 4:
    st.warning("⚠️ 请上传所有必要文件后继续")
    st.stop()
else:
    st.success("✅ 文件上传完成")

# -------------------------
# 工具函数
# -------------------------
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

def find_sheet(xls, keyword):
    for s in xls.sheet_names:
        if keyword in s:
            return s
    raise ValueError(f"未找到包含关键词「{keyword}」的sheet")

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

    # 日期比对
    if "日期" in main_kw or "日期" in ref_kw:
        if not same_date_ymd(main_val, ref_val):
            errors = 1
    else:
        main_num = normalize_num(main_val)
        ref_num = normalize_num(ref_val)
        if isinstance(main_num,(int,float)) and isinstance(ref_num,(int,float)):
            diff = abs(main_num-ref_num)
            if diff > ignore_tol:
                errors = 1
        else:
            if str(main_num).strip()!=str(ref_val).strip():
                errors = 1

    if errors:
        excel_row = idx + 3
        col_idx = list(main_df.columns).index(main_col)+1
        ws.cell(excel_row, col_idx).fill = red_fill
    return errors

# -------------------------
# 读取文件
# -------------------------
main_file = find_file(uploaded_files, "提成项目")
ec_file = find_file(uploaded_files, "二次明细")
fk_file = find_file(uploaded_files, "放款明细")
product_file = find_file(uploaded_files, "产品台账")

# 起租提成sheet
xls_main = pd.ExcelFile(main_file)
main_sheet = find_sheet(xls_main, "起租提成")
main_df = pd.read_excel(xls_main, sheet_name=main_sheet, header=1)

# 原始数据
ec_df = pd.read_excel(ec_file)
fk_xls = pd.ExcelFile(fk_file)

# 🆕 改为自动寻找包含“提成”的sheet
fk_ticheng_sheet = find_sheet(fk_xls, "提成")
fk_df = pd.read_excel(fk_xls, sheet_name=fk_ticheng_sheet)

# 🆕 新增包含“经理”的sheet
fk_manager_sheet = find_sheet(fk_xls, "经理")
manager_df = pd.read_excel(fk_xls, sheet_name=fk_manager_sheet)

product_df = pd.read_excel(product_file)

# 合同号列
contract_col_main = find_col(main_df, "合同")
contract_col_ec = find_col(ec_df, "合同")
contract_col_fk = find_col(fk_df, "合同")
contract_col_product = find_col(product_df, "合同")
contract_col_manager = find_col(manager_df, "合同")

# -------------------------
# 映射配置（新增“年限” vs “租赁期限”）
# -------------------------
mappings = [
    ("起租日期", ["起租日_商","起租日_商"], 0),   # 二次明细 & 产品台账
    ("租赁本金", ["租赁本金"], 0),                 # 放款明细
    ("收益率", ["XIRR_商_起租"], 0.005),           # 产品台账
    ("年限", ["租赁期限"], 0),                    # 🆕 提成sheet
]

# -------------------------
# Excel输出初始化
# -------------------------
output_path = "起租提成_审核标注版.xlsx"
wb = Workbook()
ws = wb.active

for c_idx, col_name in enumerate(main_df.columns, start=1):
    ws.cell(1, c_idx, col_name)

red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

# -------------------------
# 循环检查
# -------------------------
total_errors = 0
contracts_seen = set()

progress = st.progress(0)
status_text = st.empty()
n_rows = len(main_df)

for idx, row in main_df.iterrows():
    contract_no = str(row.get(contract_col_main)).strip()
    if pd.isna(contract_no) or contract_no in ["","nan"]:
        continue
    contracts_seen.add(contract_no)
    
    # 比对字段
    for main_kw, ref_kws, tol in mappings:
        for ref_kw, ref_df, ref_contract_col in zip(
            ref_kws,
            [ec_df, product_df] if main_kw=="起租日期" else
            [fk_df] if main_kw in ["租赁本金","年限"] else
            [product_df],
            [contract_col_ec, contract_col_product] if main_kw=="起租日期" else
            [contract_col_fk] if main_kw in ["租赁本金","年限"] else
            [contract_col_product]
        ):
            total_errors += compare_and_mark(
                idx, row, main_df, main_kw, ref_df, ref_kw, ref_contract_col, ws, red_fill, tol
            )

    # 🆕 检查操作人与客户经理匹配
    op_col = find_col(fk_df, "操作人")
    mgr_col = find_col(manager_df, "客户经理")
    if op_col and mgr_col:
        op_val = fk_df.loc[fk_df[contract_col_fk].astype(str).str.strip()==contract_no, op_col]
        mgr_val = manager_df.loc[manager_df[contract_col_manager].astype(str).str.strip()==contract_no, mgr_col]
        if not op_val.empty and not mgr_val.empty:
            if str(op_val.iloc[0]).strip() != str(mgr_val.iloc[0]).strip():
                excel_row = idx + 3
                op_main_col = find_col(main_df, "操作人")
                if op_main_col:
                    col_idx = list(main_df.columns).index(op_main_col) + 1
                    ws.cell(excel_row, col_idx).fill = red_fill
                total_errors += 1

    progress.progress((idx+1)/n_rows)
    if (idx+1) % 10 == 0 or idx+1 == n_rows:
        status_text.text(f"正在检查... {idx+1}/{n_rows} 行")

# -------------------------
# 标黄合同号列 & 写入数据
# -------------------------
contract_col_idx_excel = list(main_df.columns).index(contract_col_main)+1
for row_idx in range(len(main_df)):
    excel_row = row_idx+3
    has_red = any(ws.cell(excel_row,c).fill==red_fill for c in range(1,len(main_df.columns)+1))
    if has_red:
        ws.cell(excel_row, contract_col_idx_excel).fill = yellow_fill
    # 写入原数据
    for c_idx, val in enumerate(main_df.iloc[row_idx], start=1):
        ws.cell(excel_row, c_idx, val)

# -------------------------
# 导出Excel
# -------------------------
output_stream = BytesIO()
wb.save(output_stream)
output_stream.seek(0)

st.download_button(
    label="📥 下载起租提成审核标注版",
    data=output_stream,
    file_name="起租提成_审核标注版.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.success(f"✅ 审核完成，共发现 {total_errors} 处错误")
