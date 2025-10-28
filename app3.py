# =====================================
# Streamlit App: 人事用“提成项目 & 二次项目 & 平台工 & 独立架构 & 低价值”自动审核（扩展版）
# - 严格控制字段比对逻辑
# - 日期解析容错
# - “租赁期限”±0.5 月误差（经理表年 -> 乘12）
# - ✅ 操作人 vs 客户经理
# - ✅ 产品 vs 产品名称_商
# - ✅ 城市经理 vs 超期明细 城市经理
# - ✅ 忽略空合同号、大小写差异、全角/半角差异
# =====================================

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import unicodedata, re

st.title("📊 人事用审核工具（扩展+城市经理校验）")

# ========== 上传文件 ==========
uploaded_files = st.file_uploader(
    "请上传原始数据表（提成项目、二次明细、放款明细、产品台账、超期明细）",
    type="xlsx", accept_multiple_files=True
)

if not uploaded_files or len(uploaded_files) < 5:
    st.warning("⚠️ 请上传至少五个文件（提成项目、二次明细、放款明细、产品台账、超期明细）")
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
    if df is None:
        return None
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
            s = s.replace("%", "")
            return float(s) / 100
        return float(s)
    except:
        return s

def normalize_text(val):
    """文本清洗：去除空格、全角、大小写"""
    if pd.isna(val):
        return ""
    s = str(val)
    s = re.sub(r'[\n\r\t ]+', '', s)
    s = s.replace('\u3000', '')  # 全角空格
    s = ''.join(unicodedata.normalize('NFKC', ch) for ch in s)
    return s.lower().strip()

def detect_header_row(file, sheet_name):
    preview = pd.read_excel(file, sheet_name=sheet_name, nrows=2, header=None)
    first_row = preview.iloc[0]
    total_cells = len(first_row)
    empty_like = sum(
        (pd.isna(x) or str(x).startswith("Unnamed") or str(x).strip() == "")
        for x in first_row
    )
    empty_ratio = empty_like / total_cells if total_cells > 0 else 0
    return 1 if empty_ratio >= 0.7 else 0

def get_header_row(file, sheet_name):
    if any(k in sheet_name for k in ["起租", "二次"]):
        return 1
    return detect_header_row(file, sheet_name)

# ========== 比对函数 ==========
def compare_and_mark(
    idx, row, main_df, main_kw, ref_df, ref_kw, ref_contract_col,
    ws, red_fill, contract_col_main, ignore_tol=0, multiplier=1
):
    main_col = find_col(main_df, main_kw)
    ref_col = find_col(ref_df, ref_kw)
    if not main_col or not ref_col or not ref_contract_col:
        return 0

    contract_no = str(row.get(contract_col_main)).strip()
    if pd.isna(contract_no) or contract_no in ["", "nan", "None", "none"]:
        return 0

    ref_rows = ref_df[ref_df[ref_contract_col].astype(str).str.strip() == contract_no]
    if ref_rows.empty:
        return 0

    ref_val = ref_rows.iloc[0][ref_col]
    main_val = row.get(main_col)
    if pd.isna(main_val) and pd.isna(ref_val):
        return 0

    errors = 0

    # ---- 年限 / 租赁期限 ----
    if any(k in main_kw for k in ["年限", "租赁期限"]):
        main_num = normalize_num(main_val)
        ref_num = normalize_num(ref_val)
        if isinstance(ref_num, (int, float)):
            ref_num = ref_num * multiplier
        if isinstance(main_num, (int, float)) and isinstance(ref_num, (int, float)):
            if abs(main_num - ref_num) > 0.5:
                errors = 1
        else:
            if normalize_text(main_val) != normalize_text(ref_val):
                errors = 1

    # ---- 日期字段 ----
    elif "日期" in main_kw or any(word in main_kw for word in ["起租日", "起租日期", "起租"]):
        main_dt = pd.to_datetime(main_val, errors='coerce')
        ref_dt = pd.to_datetime(ref_val, errors='coerce')
        if pd.isna(main_dt) or pd.isna(ref_dt):
            errors = 1
        else:
            if not (main_dt.year == ref_dt.year and main_dt.month == ref_dt.month and main_dt.day == ref_dt.day):
                errors = 1

    # ---- 数值 / 文本字段 ----
    else:
        main_num = normalize_num(main_val)
        ref_num = normalize_num(ref_val)
        if isinstance(main_num, (int, float)) and isinstance(ref_num, (int, float)):
            if abs(main_num - ref_num) > ignore_tol:
                errors = 1
        else:
            if normalize_text(main_val) != normalize_text(ref_val):
                errors = 1

    # ---- 标红 ----
    if errors:
        excel_row = idx + 2 + header_offset
        col_idx = list(main_df.columns).index(main_col) + 1
        ws.cell(excel_row, col_idx).fill = red_fill

    return errors

# ========== 文件读取 ==========
main_file = find_file(uploaded_files, "提成项目")
ec_file = find_file(uploaded_files, "二次明细")
fk_file = find_file(uploaded_files, "放款明细")
product_file = find_file(uploaded_files, "产品台账")
overdue_file = find_file(uploaded_files, "超期明细")

ec_df = pd.read_excel(ec_file)
fk_xls = pd.ExcelFile(fk_file)
fk_df = pd.read_excel(fk_xls, sheet_name=[s for s in fk_xls.sheet_names if "本司" in s][0])
mgr_sheets = [s for s in fk_xls.sheet_names if "经理" in s]
manager_df = pd.read_excel(fk_xls, sheet_name=mgr_sheets[0]) if mgr_sheets else None
product_df = pd.read_excel(product_file)
overdue_df = pd.read_excel(overdue_file)

contract_col_ec = find_col(ec_df, "合同")
contract_col_fk = find_col(fk_df, "合同")
contract_col_mgr = find_col(manager_df, "合同") if manager_df is not None else None
contract_col_product = find_col(product_df, "合同")
contract_col_overdue = find_col(overdue_df, "合同")

# ========== 审核函数 ==========
def audit_sheet(sheet_name, main_file, ec_df, fk_df, product_df, manager_df, overdue_df):
    xls_main = pd.ExcelFile(main_file)
    global header_offset
    header_row = get_header_row(main_file, sheet_name)
    header_offset = header_row
    main_df = pd.read_excel(xls_main, sheet_name=sheet_name, header=header_row)
    st.write(f"📘 审核中：{sheet_name}（header={header_row}）")

    contract_col_main = find_col(main_df, "合同")
    if not contract_col_main:
        st.error(f"❌ {sheet_name} 中未找到“合同”列，已跳过。")
        return None, 0

    # ==== 对照规则 ====
    mapping_rules = {
        "起租日期": [
            (ec_df, "起租日_商", contract_col_ec, 1, 0),
            (product_df, "起租日_商", contract_col_product, 1, 0),
        ],
        "租赁本金": [(fk_df, "租赁本金", contract_col_fk, 1, 0)],
        "收益率": [(product_df, "XIRR_商_起租", contract_col_product, 1, 0.005)],
        "租赁期限": [(manager_df, "租赁期限", contract_col_mgr, 12, 0)],
        "操作人": [(fk_df, "客户经理", contract_col_fk, 1, 0)],
        "客户经理": [(fk_df, "客户经理", contract_col_fk, 1, 0)],
        "产品": [(product_df, "产品名称_商", contract_col_product, 1, 0)],
        # ✅ 新增：城市经理校验
        "城市经理": [(overdue_df, "城市经理", contract_col_overdue, 1, 0)]
    }

    wb = Workbook()
    ws = wb.active
    for c_idx, col_name in enumerate(main_df.columns, start=1):
        ws.cell(1 + header_offset, c_idx, col_name)

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    total_errors = 0
    n_rows = len(main_df)
    progress = st.progress(0)
    status = st.empty()

    for idx, row in main_df.iterrows():
        for main_kw, refs in mapping_rules.items():
            actual_main_col = find_col(main_df, main_kw)
            if not actual_main_col:
                continue
            for ref_df, ref_kw, ref_contract_col, mult, tol in refs:
                if ref_df is None or ref_contract_col is None:
                    continue
                total_errors += compare_and_mark(
                    idx, row, main_df, main_kw, ref_df, ref_kw,
                    ref_contract_col, ws, red_fill,
                    contract_col_main, ignore_tol=tol, multiplier=mult
                )

        progress.progress((idx + 1) / n_rows)
        if (idx + 1) % 10 == 0 or idx + 1 == n_rows:
            status.text(f"{sheet_name}：{idx + 1}/{n_rows} 行")

    # ==== 标黄合同号列 ====
    contract_col_idx_excel = list(main_df.columns).index(contract_col_main) + 1
    for row_idx in range(len(main_df)):
        excel_row = row_idx + 2 + header_offset
        has_red = any(ws.cell(excel_row, c).fill == red_fill for c in range(1, len(main_df.columns) + 1))
        if has_red:
            ws.cell(excel_row, contract_col_idx_excel).fill = yellow_fill
        for c_idx, val in enumerate(main_df.iloc[row_idx], start=1):
            ws.cell(excel_row, c_idx, val)

    # ==== 导出 ====
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

# ========== 执行 ==========
xls_main = pd.ExcelFile(main_file)
target_sheets = [s for s in xls_main.sheet_names if any(k in s for k in ["起租", "二次", "平台工", "独立架构", "低价值"])]

if not target_sheets:
    st.warning("⚠️ 未找到目标 sheet。")
else:
    for sheet_name in target_sheets:
        audit_sheet(sheet_name, main_file, ec_df, fk_df, product_df, manager_df, overdue_df)
