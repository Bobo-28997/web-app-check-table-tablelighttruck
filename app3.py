# =====================================
# Streamlit App: 人事用“提成项目 & 二次项目 & 平台工 & 独立架构 & 低价值 & 权责发生”自动审核（终极修正版）
# - 严格字段比对
# - 日期容错
# - “租赁期限”±0.5 月（经理表年 -> ×12）
# - ✅ 操作人 vs 客户经理
# - ✅ 产品 vs 产品名称_商
# - ✅ 城市经理 vs 超期明细 城市经理
# - ✅ 权责发生字段 vs 经理表字段
# - ✅ 最终漏填检测：使用“放款明细”中含“提成”的sheet为基准
# =====================================

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows # <--- 添加这一行
from io import BytesIO
import unicodedata, re

st.title("📊 模拟人事用薪资计算表自动审核系统-2")

# ========== 上传文件 ==========
uploaded_files = st.file_uploader(
    "请上传文件名中包含以下字段的文件：提成项目、放款明细、二次明细、产品台账，超期明细。最后誊写，需检的表为提成项目表。使用2数据时，提成项目表中需检sheet的‘年限’等月单位的时长列名需改为‘租赁期限’。",
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
    if pd.isna(val):
        return ""
    s = str(val)
    s = re.sub(r'[\n\r\t ]+', '', s)
    s = s.replace('\u3000', '')
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

def normalize_contract_key(series: pd.Series) -> pd.Series:
    """
    对合同号 Series 进行标准化处理，用于安全的 pd.merge 操作。
    (来自我们上一个 app 的经验)
    """
    s = series.astype(str)
    s = s.str.replace(r"\.0$", "", regex=True) 
    s = s.str.strip()
    s = s.str.upper() 
    s = s.str.replace('－', '-', regex=False)
    # 注意：这里我们不用 normalize_text 的 r'\s+'
    # 因为合同号内部可能允许有空格
    return s

def prepare_one_ref_df(ref_df, ref_contract_col, required_cols, prefix):
    """
    预处理单个参考DataFrame，提取所需列并标准化Key。
    (V2: 增加未找到列的警告)
    """
    if ref_df is None:
        st.warning(f"⚠️ 参考文件 '{prefix}' 未加载 (df is None)。")
        return pd.DataFrame(columns=['__KEY__'])
        
    if ref_contract_col is None:
        st.warning(f"⚠️ 在 {prefix} 参考表中未找到'合同'列，跳过此数据源。")
        return pd.DataFrame(columns=['__KEY__'])

    # 找出实际存在的列
    cols_to_extract = []
    col_mapping = {} # '原始列名' -> 'ref_prefix_原始列名'

    for col_kw in required_cols:
        actual_col = find_col(ref_df, col_kw)
        
        if actual_col:
            cols_to_extract.append(actual_col)
            # 使用原始列名 (col_kw) 作为标准后缀
            col_mapping[actual_col] = f"ref_{prefix}_{col_kw}"
        else:
            # --- VVVV (【新功能】添加警告) VVVV ---
            st.warning(f"⚠️ 在 {prefix} 参考表中未找到列 (关键字: '{col_kw}')")
            # --- ^^^^ (【新功能】添加警告) ^^^^ ---
            
    if not cols_to_extract:
        # 如果一个需要的列都没找到，也不用继续了
        st.warning(f"⚠️ 在 {prefix} 参考表中未找到任何所需字段，跳过。")
        return pd.DataFrame(columns=['__KEY__'])

    # 提取所需列 + 合同列
    cols_to_extract.append(ref_contract_col)
    
    # 确保所有列都是唯一的
    cols_to_extract_unique = list(set(cols_to_extract))
    
    # 检查所有列是否真的存在 (set 操作可能引入不存在的列, 尽管这里不太可能)
    valid_cols = [col for col in cols_to_extract_unique if col in ref_df.columns]
    
    std_df = ref_df[valid_cols].copy()

    # 标准化Key
    std_df['__KEY__'] = normalize_contract_key(std_df[ref_contract_col])
    
    # 重命名
    std_df = std_df.rename(columns=col_mapping)
    
    # 只保留需要的列
    final_cols = ['__KEY__'] + list(col_mapping.values())
    
    # 确保 final_cols 都在 std_df 中 (因为重命名了)
    final_cols_in_df = [col for col in final_cols if col in std_df.columns]
    
    std_df = std_df[final_cols_in_df]
    
    # 去重
    std_df = std_df.drop_duplicates(subset=['__KEY__'], keep='first')
    return std_df

def compare_series_vec(s_main, s_ref, compare_type='text', tolerance=0, multiplier=1):
    """
    向量化比较两个Series，复刻 compare_and_mark 的逻辑。
    (V2：增加对 merge 失败 (NaN) 的静默跳过)
    """
    # 0. 识别 Merge 失败
    merge_failed_mask = s_ref.isna()

    # 1. 预处理空值
    main_is_na = pd.isna(s_main) | (s_main.astype(str).str.strip().isin(["", "nan", "None"]))
    ref_is_na = pd.isna(s_ref) | (s_ref.astype(str).str.strip().isin(["", "nan", "None"]))
    both_are_na = main_is_na & ref_is_na
    
    errors = pd.Series(False, index=s_main.index)

    # 2. 日期比较
    if compare_type == 'date':
        d_main = pd.to_datetime(s_main, errors='coerce')
        d_ref = pd.to_datetime(s_ref, errors='coerce')
        
        # 仅在两者都是有效日期时比较
        valid_dates_mask = d_main.notna() & d_ref.notna()
        date_diff_mask = (d_main.dt.date != d_ref.dt.date)
        errors = valid_dates_mask & date_diff_mask
        
        # 如果一个是日期，另一个不是（且不为空），也算错误
        one_is_date_one_is_not = (d_main.notna() & d_ref.isna() & ~ref_is_na) | \
                                 (d_main.isna() & ~main_is_na & d_ref.notna())
        errors |= one_is_date_one_is_not

    # 3. 数值比较 (包括特殊的租赁期限)
    elif compare_type == 'num':
        s_main_norm = s_main.apply(normalize_num)
        s_ref_norm = s_ref.apply(normalize_num)
        
        # 应用乘数
        if multiplier != 1:
            s_ref_norm = pd.to_numeric(s_ref_norm, errors='coerce') * multiplier
        
        # 检查是否都为数值
        is_num_main = s_main_norm.apply(lambda x: isinstance(x, (int, float)))
        is_num_ref = s_ref_norm.apply(lambda x: isinstance(x, (int, float)))
        both_are_num = is_num_main & is_num_ref

        if both_are_num.any():
            diff = (s_main_norm[both_are_num] - s_ref_norm[both_are_num]).abs()
            errors.loc[both_are_num] = (diff > (tolerance + 1e-6)) # 1e-6 避免浮点精度问题
            
        # 如果一个是数字，另一个是文本（且不为空），也算错误
        one_is_num_one_is_not = (is_num_main & ~is_num_ref & ~ref_is_na) | \
                                (~is_num_main & ~main_is_na & is_num_ref)
        errors |= one_is_num_one_is_not

    # 4. 文本比较
    else: # compare_type == 'text'
        s_main_norm_text = s_main.apply(normalize_text)
        s_ref_norm_text = s_ref.apply(normalize_text)
        errors = (s_main_norm_text != s_ref_norm_text)

    # 5. 最终错误逻辑
    final_errors = errors & ~both_are_na
    
    # 排除 "Merge 失败" 导致的错误 (复刻 'if ref_rows.empty: return 0')
    lookup_failure_mask = merge_failed_mask & ~main_is_na
    final_errors = final_errors & ~lookup_failure_mask
    
    return final_errors

# ========== 比对函数 ==========
# =====================================
# 🧮 审核函数 (向量化版)
# =====================================
def audit_sheet_vec(sheet_name, main_file, all_std_dfs, mapping_rules_vec):
    xls_main = pd.ExcelFile(main_file)
    
    # 1. 读取主表 (尊重动态表头)
    header_offset = get_header_row(main_file, sheet_name)
    main_df = pd.read_excel(xls_main, sheet_name=sheet_name, header=header_offset)
    st.write(f"📘 审核中：{sheet_name}（header={header_offset}）")

    contract_col_main = find_col(main_df, "合同")
    if not contract_col_main:
        st.error(f"❌ {sheet_name} 中未找到“合同”列，已跳过。")
        return None, 0

    # 2. 准备主表
    main_df['__ROW_IDX__'] = main_df.index
    main_df['__KEY__'] = normalize_contract_key(main_df[contract_col_main])

    # 3. 一次性合并所有参考数据
    merged_df = main_df.copy()
    for std_df in all_std_dfs.values():
        if not std_df.empty:
            merged_df = pd.merge(merged_df, std_df, on='__KEY__', how='left')

    # 4. === 遍历字段进行向量化比对 ===
    total_errors = 0
    errors_locations = set() # 存储 (row_idx, col_name)
    row_has_error = pd.Series(False, index=merged_df.index)

    progress = st.progress(0)
    status = st.empty()
    
    total_comparisons = len(mapping_rules_vec)
    current_comparison = 0

    for main_kw, comparisons in mapping_rules_vec.items():
        current_comparison += 1
        
        main_col = find_col(main_df, main_kw)
        if not main_col:
            continue # 跳过主表中不存在的列
        
        status.text(f"检查「{sheet_name}」: {main_kw}...")
        
        # 存储此字段的最终错误
        field_error_mask = pd.Series(False, index=merged_df.index)
        
        for (ref_col, compare_type, tol, mult) in comparisons:
            if ref_col not in merged_df.columns:
                continue # 跳过 merge 失败或未定义的参考列
            
            s_main = merged_df[main_col]
            s_ref = merged_df[ref_col]
            
            # 获取此单一比对的错误
            # (注意：如果一个字段有多个比对源, 它们是 'OR' 关系)
            # (即, 只要和 *一个* 源匹配成功, 就不算错... 
            #  ...等一下, 原始逻辑是 (err1 + err2 + ...), 
            #  这意味着只要 *一个* 源 *不* 匹配, 就算错)
            
            errors_mask = compare_series_vec(s_main, s_ref, compare_type, tol, mult)
            
            # 累加错误 (原始逻辑是 total_errors +=, 意味着一个错就算错)
            field_error_mask |= errors_mask
        
        if field_error_mask.any():
            total_errors += field_error_mask.sum()
            row_has_error |= field_error_mask
            
            # 存储错误位置 (使用 __ROW_IDX__ 和 原始 main_col 名称)
            bad_indices = merged_df[field_error_mask]['__ROW_IDX__']
            for idx in bad_indices:
                errors_locations.add((idx, main_col))
        
        progress.progress(current_comparison / total_comparisons)

    status.text(f"「{sheet_name}」比对完成，正在生成标注文件...")

    # 5. === 快速写入 Excel 并标注 ===
    wb = Workbook()
    ws = wb.active
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

   # c. 准备坐标映射 (我们把 c 移到 b 之前)
    original_cols_list = list(main_df.drop(columns=['__ROW_IDX__', '__KEY__']).columns)
    col_name_to_idx = {name: i + 1 for i, name in enumerate(original_cols_list)}
    
    # a. 写入表头前的空行 (如果需要)
    if header_offset > 0:
        for _ in range(header_offset):
            # (修正：使用 original_cols_list 的长度, 而不是 main_df.columns 的长度)
            ws.append([""] * len(original_cols_list)) # 添加空行
            
    # b. 使用 dataframe_to_rows 快速写入表头 + 数据
    #    (注意：我们在这里传入了 original_cols_list, 确保列序正确)
    for r in dataframe_to_rows(main_df[original_cols_list], index=False, header=True):
        ws.append(r)

    # d. 标红错误单元格
    for (row_idx, col_name) in errors_locations:
        if col_name in col_name_to_idx:
            excel_row = row_idx + 1 + header_offset + 1 # (row_idx 0-based) + (1 for header) + (offset) + (1 for 1-based)
            excel_col = col_name_to_idx[col_name]
            ws.cell(excel_row, excel_col).fill = red_fill
            
    # e. 标黄有错误的合同号
    if contract_col_main in col_name_to_idx:
        contract_col_excel_idx = col_name_to_idx[contract_col_main]
        error_row_indices = merged_df[row_has_error]['__ROW_IDX__']
        for row_idx in error_row_indices:
            excel_row = row_idx + 1 + header_offset + 1
            ws.cell(excel_row, contract_col_excel_idx).fill = yellow_fill

    # 6. 导出
    output_stream = BytesIO()
    wb.save(output_stream)
    output_stream.seek(0)
    st.download_button(
        label=f"📥 下载 {sheet_name} 审核标注版",
        data=output_stream,
        file_name=f"{sheet_name}_审核标注版.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"download_{sheet_name}" # 添加唯一的key
    )

    # --- VVVV (【新功能】从这里开始添加) VVVV ---

    # 7. (新) 导出仅含错误行的文件 (带标红)
    if row_has_error.any():
        try:
            # 1. 获取仅含错误行的 DataFrame (只保留原始列)
            df_errors_only = merged_df.loc[row_has_error, original_cols_list].copy()
            
            # 2. 关键：创建 "原始行索引" 到 "新Excel行号" 的映射
            #    我们获取所有出错行的 __ROW_IDX__
            original_indices_with_error = merged_df.loc[row_has_error, '__ROW_IDX__']
            
            #    创建映射: { 原始索引 : 新的Excel行号 }
            #    (enumerate start=2, 因为 Excel 行 1 是表头, 数据从行 2 开始)
            original_idx_to_new_excel_row = {
                original_idx: new_row_num 
                for new_row_num, original_idx in enumerate(original_indices_with_error, start=2)
            }

            # 3. 创建一个新的工作簿(Workbook)
            wb_errors = Workbook()
            ws_errors = wb_errors.active
            
            # 4. 使用 dataframe_to_rows 快速写入数据
            for r in dataframe_to_rows(df_errors_only, index=False, header=True):
                ws_errors.append(r)
                
            # 5. 遍历主错误列表(errors_locations)，进行标红
            #    (col_name_to_idx 和 red_fill 已在前面定义)
            for (original_row_idx, col_name) in errors_locations:
                
                # 检查这个错误是否在我们 "仅错误行" 的映射中
                if original_row_idx in original_idx_to_new_excel_row:
                    
                    # 获取它在新Excel文件中的行号
                    new_row = original_idx_to_new_excel_row[original_row_idx]
                    
                    # 获取列号
                    if col_name in col_name_to_idx:
                        new_col = col_name_to_idx[col_name]
                        
                        # 应用标红
                        ws_errors.cell(row=new_row, column=new_col).fill = red_fill
            
            # 6. 保存到 BytesIO
            output_errors_only = BytesIO()
            wb_errors.save(output_errors_only)
            output_errors_only.seek(0)
            
            # 7. 创建第二个下载按钮
            st.download_button(
                label=f"📥 下载 {sheet_name} (仅含错误行, 带标红)", # 更新了标签
                data=output_errors_only,
                file_name=f"{sheet_name}_仅错误行_标红.xlsx", # 更新了文件名
                key=f"download_{sheet_name}_errors_only" # 必须使用唯一的 key
            )
        except Exception as e:
            st.error(f"❌ 生成“仅错误行”文件时出错: {e}")
            
    # --- ^^^^ (【新功能】到这里结束) ^^^^ ---
    st.success(f"✅ {sheet_name} 审核完成，共发现 {total_errors} 处错误")
    
    # 返回原始的 main_df (不含 __KEY__), 用于漏填检测
    return main_df.drop(columns=['__ROW_IDX__', '__KEY__']), total_errors

# ========== 文件读取 & 预处理 ==========
main_file = find_file(uploaded_files, "提成项目")
ec_file = find_file(uploaded_files, "二次明细")
fk_file = find_file(uploaded_files, "放款明细")
product_file = find_file(uploaded_files, "产品台账")
overdue_file = find_file(uploaded_files, "超期明细")

st.info("ℹ️ 正在读取并预处理参考文件...")

# 1. 加载所有参考 DF
ec_df = pd.read_excel(ec_file)
fk_xls = pd.ExcelFile(fk_file)
fk_df = pd.read_excel(fk_xls, sheet_name=[s for s in fk_xls.sheet_names if "本司" in s][0])
product_df = pd.read_excel(product_file)
overdue_df = pd.read_excel(overdue_file)

# ---- 新增提成sheet提取 ----
commission_sheets = [s for s in fk_xls.sheet_names if "提成" in s]
commission_df = pd.read_excel(fk_xls, sheet_name=commission_sheets[0]) if commission_sheets else None

# ---- 找到所有参考表的合同列 ----
contract_col_ec = find_col(ec_df, "合同")
contract_col_fk = find_col(fk_df, "合同")
contract_col_comm = find_col(commission_df, "合同") if commission_df is not None else None
contract_col_product = find_col(product_df, "合同")
contract_col_overdue = find_col(overdue_df, "合同")

# 2. (新) 定义向量化映射规则
# 格式: { "主表列名": [ (参考列表名, 比较类型, 容差, 乘数), ... ] }
mapping_rules_vec = {
    "起租日期": [
        ("ref_ec_起租日_商", 'date', 0, 1),
        ("ref_product_起租日_商", 'date', 0, 1)
    ],
    "租赁本金": [("ref_fk_租赁本金", 'num', 0, 1)],
    "收益率": [("ref_product_XIRR_商_起租", 'num', 0.005, 1)],
    "操作人": [("ref_fk_客户经理", 'text', 0, 1)],
    "客户经理": [("ref_fk_客户经理", 'text', 0, 1)],
    "产品": [("ref_product_产品名称_商", 'text', 0, 1)],
    "城市经理": [("ref_overdue_城市经理", 'text', 0, 1)],
}

# 3. (新) 预处理所有参考 DF
# 从 mapping_rules_vec 中提取所有需要的列
ec_cols = ["起租日_商"]
fk_cols = ["租赁本金", "客户经理"]
product_cols = ["起租日_商", "XIRR_商_起租", "产品名称_商"]
overdue_cols = ["城市经理"]

ec_std = prepare_one_ref_df(ec_df, contract_col_ec, ec_cols, "ec")
fk_std = prepare_one_ref_df(fk_df, contract_col_fk, fk_cols, "fk")
product_std = prepare_one_ref_df(product_df, contract_col_product, product_cols, "product")
overdue_std = prepare_one_ref_df(overdue_df, contract_col_overdue, overdue_cols, "overdue")

all_std_dfs = {
    "ec": ec_std,
    "fk": fk_std,
    "product": product_std,
    "overdue": overdue_std
}

st.success("✅ 参考文件预处理完成。")

# ========== 执行主流程 (向量化) ==========
xls_main = pd.ExcelFile(main_file)
target_sheets = [
    s for s in xls_main.sheet_names
    if any(k in s for k in ["起租", "二次", "平台工", "独立架构", "低价值", "权责发生"])
]

all_contracts_in_sheets = set()

if not target_sheets:
    st.warning("⚠️ 未找到目标 sheet。")
else:
    for sheet_name in target_sheets:
        # (新) 调用向量化审核函数
        df, _ = audit_sheet_vec(sheet_name, main_file, all_std_dfs, mapping_rules_vec)
        
        if df is not None:
            col = find_col(df, "合同")
            if col:
                # (新) 标准化合同号, 用于 set.update
                normalized_contracts = normalize_contract_key(df[col].dropna())
                all_contracts_in_sheets.update(normalized_contracts)

# ======= 新逻辑：使用“提成”sheet合同号检测漏填 =======
if commission_df is not None and contract_col_comm:
    # (新) 必须同样标准化提成表的合同号
    commission_contracts = set(normalize_contract_key(commission_df[contract_col_comm].dropna()))
    
    missing_contracts = sorted(list(commission_contracts - all_contracts_in_sheets))

    # --- VVVV (从这里开始, 修复了缩进) VVVV ---
    st.subheader("📋 合同漏填检测结果（基于提成sheet）")
    st.write(f"共 {len(missing_contracts)} 个合同在六张表中未出现。")

    if missing_contracts:
        wb_miss = Workbook()
        ws_miss = wb_miss.active
        ws_miss.cell(1, 1, "未出现在任一表中的合同号")
        yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        for i, cno in enumerate(missing_contracts, start=2):
            ws_miss.cell(i, 1, cno).fill = yellow

        out_miss = BytesIO()
        wb_miss.save(out_miss)
        out_miss.seek(0)
        st.download_button(
            "📥 下载漏填合同列表",
            data=out_miss,
            file_name="漏填合同号列表.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("✅ 所有提成sheet合同号均已出现在六张表中，无漏填。")
_ # --- ^^^^ (到这里结束) ^^^^ ---
