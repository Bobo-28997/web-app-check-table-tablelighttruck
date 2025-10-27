# =====================================
# Streamlit App: 人事用“提成项目 & 二次项目 & 平台工”自动审核（改进版）
# - 严格控制各字段对照表
# - 日期解析更稳健（只在两端都能解析为日期时比较；否则视为不一致）
# - 仅“年限/租赁期限”允许 ±0.5 月误差（经理表年 -> 乘12）
# =====================================

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from io import BytesIO

st.title("📊 人事用审核工具（改进）：起租/二次/平台工 + 经理年限比对")

# ========== 上传文件 ==========
uploaded_files = st.file_uploader(
    "请上传原始数据表（提成项目、二次明细、放款明细、产品台账等）",
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
            s = s.replace("%","")
            return float(s)/100
        return float(s)
    except:
        return s

def same_date_ymd(a,b):
    """
    比较年月日，忽略时分秒。
    返回 True 当且仅当两端都能解析为日期并且年月日相同。
    """
    try:
        da = pd.to_datetime(a, errors='coerce')
        db = pd.to_datetime(b, errors='coerce')
        if pd.isna(da) or pd.isna(db):
            return False
        return (da.year==db.year) and (da.month==db.month) and (da.day==db.day)
    except:
        return False

def detect_header_row(file, sheet_name):
    preview = pd.read_excel(file, sheet_name=sheet_name, nrows=2, header=None)
    first_row = preview.iloc[0]
    total_cells = len(first_row)
    empty_like = sum(
        (pd.isna(x) or str(x).startswith("Unnamed") or str(x).strip() == "")
        for x in first_row
    )
    empty_ratio = empty_like / total_cells if total_cells > 0 else 0
    if empty_ratio >= 0.7:
        return 1
    return 0

def get_header_row(file, sheet_name):
    # 起租/二次通常 header 在第2行（保留白名单）
    if any(k in sheet_name for k in ["起租", "二次"]):
        return 1
    return detect_header_row(file, sheet_name)

# ========== compare_and_mark（改进版） ==========
def compare_and_mark(
    idx, row, main_df, main_kw, ref_df, ref_kw, ref_contract_col,
    ws, red_fill, contract_col_main, ignore_tol=0, multiplier=1
):
    """
    - main_kw: 主表关键字（例如 '起租日期' / '年限' / '租赁本金' / '收益率'）
    - ref_df/ref_kw: 参考表和对应列名关键词
    - multiplier: 当参考值需要换算时使用（经理表年 -> 乘12）
    - ignore_tol: 仅用于普通数值字段的容差（租赁期限有自己处理）
    """
    # 基本列存在性检查
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

    # 两端都为空 -> 无差异
    if pd.isna(main_val) and pd.isna(ref_val):
        return 0

    errors = 0

    # ---- 1) 年限 / 租赁期限 专属处理 ----
    if any(k in main_kw for k in ["年限", "租赁期限"]):
        main_num = normalize_num(main_val)
        ref_num = normalize_num(ref_val)
        if isinstance(ref_num, (int, float)):
            ref_num = ref_num * multiplier  # 经理表是年 -> 转月
        # 若两端均为数值则按月比较，允许0.5月误差
        if isinstance(main_num, (int, float)) and isinstance(ref_num, (int, float)):
            if abs(main_num - ref_num) > 0.5:
                errors = 1
        else:
            # 如果任何一端不是数值（比如字符串），视为不匹配
            if str(main_val).strip() != str(ref_val).strip():
                errors = 1

    # ---- 2) 日期类字段（严格） ----
    elif "日期" in main_kw or "日期" in ref_kw or any(word in main_kw for word in ["起租日","起租日期","起租"]):
        # 仅在两端都能解析为日期时按年月日比较；否则判为不一致（即标红）
        main_dt = pd.to_datetime(main_val, errors='coerce')
        ref_dt = pd.to_datetime(ref_val, errors='coerce')
        if pd.isna(main_dt) or pd.isna(ref_dt):
            # 如果至少一端无法解析成日期，则认为不一致（这样产品台账上非日期数字会被标错）
            errors = 1
        else:
            if not (main_dt.year == ref_dt.year and main_dt.month == ref_dt.month and main_dt.day == ref_dt.day):
                errors = 1

    # ---- 3) 其余数值/文本字段 ----
    else:
        main_num = normalize_num(main_val)
        ref_num = normalize_num(ref_val)
        if isinstance(main_num, (int, float)) and isinstance(ref_num, (int, float)):
            if abs(main_num - ref_num) > ignore_tol:
                errors = 1
        else:
            if str(main_num).strip() != str(ref_num).strip():
                errors = 1

    # 标红
    if errors:
        excel_row = idx + 2 + header_offset
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

# 放款本司表与经理表
fk_df = pd.read_excel(fk_xls, sheet_name=[s for s in fk_xls.sheet_names if "本司" in s][0])
# 找包含"经理"字样的sheet（若无则设为None）
mgr_sheets = [s for s in fk_xls.sheet_names if "经理" in s]
manager_df = pd.read_excel(fk_xls, sheet_name=mgr_sheets[0]) if mgr_sheets else None

product_df = pd.read_excel(product_file)

# 参考表合同列名识别
contract_col_ec = find_col(ec_df, "合同")
contract_col_fk = find_col(fk_df, "合同")
contract_col_mgr = find_col(manager_df, "合同") if manager_df is not None else None
contract_col_product = find_col(product_df, "合同")

# ========== 审核函数（每个sheet独立） ==========
def audit_sheet(sheet_name, main_file, ec_df, fk_df, product_df, manager_df):
    xls_main = pd.ExcelFile(main_file)
    global header_offset
    header_row = get_header_row(main_file, sheet_name)
    header_offset = header_row
    main_df = pd.read_excel(xls_main, sheet_name=sheet_name, header=header_row)
    st.write(f"📘 正在审核：{sheet_name}（header={header_row}）")

    contract_col_main = find_col(main_df, "合同")
    if not contract_col_main:
        st.error(f"❌ {sheet_name} 中未找到“合同”列，已跳过。")
        return None, 0

    # 映射：主字段 -> 要对照的 (ref_df, ref_kw, ref_contract_col, multiplier, tol)
    # 明确写出每一对，避免错误配对
    mapping_rules = {
        "起租日期": [
            (ec_df, "起租日_商", contract_col_ec, 1, 0),
            (product_df, "起租日_商", contract_col_product, 1, 0),
        ],
        "租赁本金": [
            (fk_df, "租赁本金", contract_col_fk, 1, 0),
        ],
        "收益率": [
            (product_df, "XIRR_商_起租", contract_col_product, 1, 0.005),
        ],
        # 经理表年 -> 乘12
        "租赁期限": [
            (manager_df, "租赁期限", contract_col_mgr, 12, 0),
        ]
    }

    # 若主表使用“年限”列名替代“租赁期限”，我们先尝试找到哪个存在
    possible_main_year_cols = [c for c in main_df.columns if any(k in normalize_colname(c) for k in ["年限","租赁期限"])]
    if possible_main_year_cols:
        # ensure mapping keys present (we already have mapping_rules for both)
        pass

    wb = Workbook()
    ws = wb.active
    # 将列名写入输出，放在第 1 + header_offset 行（与原始表头对齐）
    for c_idx, col_name in enumerate(main_df.columns, start=1):
        ws.cell(1 + header_offset, c_idx, col_name)

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    total_errors = 0
    n_rows = len(main_df)
    progress = st.progress(0)
    status = st.empty()

    for idx, row in main_df.iterrows():
        # 针对每个主字段，按 mapping_rules 明确对照
        for main_kw, refs in mapping_rules.items():
            # 允许主表列名为“年限”或“租赁期限”两者其中之一
            # Find actual main column name that contains main_kw substring
            actual_main_col = find_col(main_df, main_kw)
            if not actual_main_col:
                continue

            for ref_df, ref_kw, ref_contract_col, mult, tol in refs:
                # 若参考表不存在（如 manager_df 可能为 None），跳过
                if ref_df is None or ref_contract_col is None:
                    continue

                total_errors += compare_and_mark(
                    idx, row, main_df, main_kw, ref_df, ref_kw, ref_contract_col,
                    ws, red_fill, contract_col_main, ignore_tol=tol, multiplier=mult
                )

        progress.progress((idx+1)/n_rows)
        if (idx+1)%10==0 or idx+1==n_rows:
            status.text(f"{sheet_name}：{idx+1}/{n_rows} 行")

    # 标黄合同号列 & 写入数据（按 header_offset 对齐）
    contract_col_idx_excel = list(main_df.columns).index(contract_col_main) + 1
    for row_idx in range(len(main_df)):
        excel_row = row_idx + 2 + header_offset
        has_red = any(ws.cell(excel_row, c).fill == red_fill for c in range(1, len(main_df.columns)+1))
        if has_red:
            ws.cell(excel_row, contract_col_idx_excel).fill = yellow_fill
        for c_idx, val in enumerate(main_df.iloc[row_idx], start=1):
            ws.cell(excel_row, c_idx, val)

    # 导出并提供下载
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

# ========== 执行审核（对包含关键字的 sheet） ==========
xls_main = pd.ExcelFile(main_file)
target_sheets = [s for s in xls_main.sheet_names if any(k in s for k in ["起租","二次","平台工"])]

if not target_sheets:
    st.warning("⚠️ 未找到目标 sheet。")
else:
    for sheet_name in target_sheets:
        audit_sheet(sheet_name, main_file, ec_df, fk_df, product_df, manager_df)


