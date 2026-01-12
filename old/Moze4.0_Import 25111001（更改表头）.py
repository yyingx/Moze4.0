import pandas as pd
import os
import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog
import openpyxl 
import csv       
from pathlib import Path 
import traceback 
import re 
import numpy as np 

# --- ANSI 颜色代码 ---
class bcolors:
    OKGREEN = '\033[92m'
    WARNING = '\033[91m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'

# --- 常量定义 ---
RULE_BOOK_PATH = Path(r"E:\天之逸2025\Moze4.0\Moze Dict.xlsx")
TARGET_DIR = Path(r"E:\天之逸2025\Moze4.0\Moze4.0_Import")

ALIPAY_MAPPING = {
    # --- 用户原有的映射 (兼容旧版或特定格式) ---
    '交易时间': '交易时间', 
    '交易分类': '交易类型', 
    '交易对方': '交易对方',
    '商品说明': '商品', 
    '收/支': '收/支', 
    '收/付款方式': '支付方式',
    '交易状态': '当前状态', 
    '交易订单号': '交易单号', 
    '商家订单号': '商户单号',
    '备注': '备注',

    # --- 新增：兼容标准/现代支付宝导出格式 ---
    '付款时间': '交易时间',     # (推荐) 使用付款时间作为标准
    '交易创建时间': '交易时间', # (备用)
    '商品名称': '商品',         # 兼容 '商品名称'
    '类型': '交易类型',       # 兼容 '类型'
}

TARGET_HEADERS = [
    "账户", "币种", "记录类型", "主类别", "子类别", "金额",  
    "手续费", "折扣", "名称", "商家", "日期", "时间",  
    "项目", "描述", "标签", "对象"
]

# --- 辅助函数 ---

def robust_date_converter(x):
    if isinstance(x, (datetime.datetime, datetime.date)):
        return pd.to_datetime(x)
    elif isinstance(x, (int, float)):
        try:
            return pd.to_datetime(x, unit='D', origin='1899-12-30') # Excel 日期 origin
        except Exception:
            return pd.NaT
    elif isinstance(x, str):
        return pd.to_datetime(x, errors='coerce')
    else:
        return pd.NaT

def get_user_inputs():
    root = tk.Tk()
    root.withdraw() 

    print("第1步: 请选择您的账单文件 (可多选 微信/支付宝)...", flush=True)
    source_files_tuple = filedialog.askopenfilenames(
        title="第1步: (可多选) 请选择 微信/支付宝 账单 (不要选合并后的)",
        filetypes=(("Excel/CSV 文件", "*.xlsx;*.csv"), ("所有文件", "*.*"))
    )
    
    if not source_files_tuple:
        print("操作已取消。", flush=True)
        root.destroy()
        return None, None
    
    source_files = [Path(f) for f in source_files_tuple]
    print(f"您选择了 {len(source_files)} 个文件:", flush=True)
    for f in source_files:
        print(f"  - {f.name}", flush=True)
    
    start_time_str = simpledialog.askstring(
        "第2步: 筛选条件 (按日期)",
        "请输入起始日期（包含当天及之后）\n\n"
        "**格式： yymmdd** (例如: 251001)\n\n"
        "*** 如果留空，将导入所有时间的数据 ***",
        parent=root
    )
    root.destroy() 

    start_datetime_obj = None
    if start_time_str and start_time_str.strip() != "":
        try:
            start_datetime_obj = datetime.datetime.strptime(start_time_str.strip(), "%y%m%d")
            print(f"筛选条件：将只导入 {start_datetime_obj.strftime('%Y-%m-%d')} 00:00:00 之后的数据。", flush=True)
        except ValueError:
            print(f"错误：日期格式不正确！'{start_time_str}' 不是 yymmdd 格式。", flush=True)
            print("程序已退出。", flush=True)
            return None, None
    else:
        print("提示：未输入起始时间，将导入文件中的所有数据。", flush=True)

    return source_files, start_datetime_obj

def load_rules(rule_path: Path):
    if not rule_path.exists():
        print(f"错误：找不到规则文件！ {rule_path}", flush=True)
        return None
        
    print(f"正在读取规则文件: {rule_path}", flush=True)
    try:
        df_rules = pd.read_excel(rule_path, engine='openpyxl')
        
        if '商家(old)' not in df_rules.columns:
            print(f"错误：规则文件 {rule_path} 中缺少 '商家(old)' 列！", flush=True)
            return None
            
        if 'is_regex' not in df_rules.columns:
            df_rules['is_regex'] = 0
        
        df_rules['is_regex'] = pd.to_numeric(df_rules['is_regex'], errors='coerce').fillna(0)
        df_rules['商家(old)'] = df_rules['商家(old)'].astype(str).str.strip()
        
        exact_mask = (df_rules['is_regex'] == 0)
        df_rules_exact = df_rules[exact_mask].drop_duplicates(subset=['商家(old)'], keep='last')
        df_rules_regex = df_rules[~exact_mask]
        
        df_rules_final = pd.concat([df_rules_exact, df_rules_regex]).reset_index(drop=True)
        
        print(f"成功加载 {len(df_rules_exact)} 条[唯一]精确规则 和 {len(df_rules_regex)} 条正则规则。", flush=True)
        return df_rules_final
        
    except Exception as e:
        print(f"读取规则文件时出错: {e}", flush=True)
        return None

def sniff_and_load_data(file_path: Path):
    print(f"\n正在分析并加载文件: {file_path.name}", flush=True)
    file_type = None
    skip_rows = 0
    df_source = None
    excel_workbook_object = None
    csv_encoding = None

    if file_path.suffix == '.csv':
        keyword = "支付宝支付科技有限公司"
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                for i, line in enumerate(f):
                    if i > 30: break
                    if keyword in line:
                        file_type, skip_rows, csv_encoding = "AliPay", i + 1, 'utf-8'
                        print(f"  检测到：【原始支付宝文件 (UTF-8)】, 抬头在第 {i+1} 行。", flush=True)
                        break
        except UnicodeDecodeError:
            try:
                with open(file_path, 'r', encoding='GBK', errors='ignore') as f:
                    for i, line in enumerate(f):
                        if i > 30: break
                        if keyword in line:
                            file_type, skip_rows, csv_encoding = "AliPay", i + 1, 'GBK'
                            print(f"  检测到：【原始支付宝文件 (GBK)】, 抬头在第 {i+1} 行。", flush=True)
                            break
            except Exception as e:
                print(f"  GBK 编码也失败: {e}", flush=True)

    elif file_path.suffix == '.xlsx':
        try:
            excel_workbook_object = openpyxl.load_workbook(file_path, read_only=True)
            sheet = excel_workbook_object.active
            cell_A1 = str(sheet['A1'].value)
            
            if cell_A1 == "交易时间":
                file_type, skip_rows = "Merged", 0
                print("  检测到：【合并后的账单】 (警告: 无法为合并文件添加来源标签)", flush=True)
            
            # ******************************************************
            # *** (!!!) 修正 4: 修正微信的错别字 (!!!) ***
            # ******************************************************
            #   将繁体的 "明細" 改回 简体的 "明细"
            elif "微信支付账单明细" in cell_A1: 
            # ******************************************************
                
                file_type, skip_rows = "WeChat", 16
                print("  检测到：【原始微信文件】", flush=True)
        except Exception as e:
            print(f"  分析XLSX文件失败: {e}", flush=True)
            if excel_workbook_object: excel_workbook_object.close()
            return None

    if file_type is None:
        print("  错误：无法识别的文件类型。跳过此文件。", flush=True)
        if excel_workbook_object: excel_workbook_object.close()
        return None

    print(f"  正在加载数据 (跳过 {skip_rows} 行)...", flush=True)
    try:
        io_source = excel_workbook_object if excel_workbook_object else file_path
        
        if file_type == "WeChat":
            df_source = pd.read_excel(io_source, skiprows=skip_rows, engine='openpyxl')
            df_source.columns = df_source.columns.str.strip()
            if '金额(元)' not in df_source.columns:
                raise ValueError("微信账单(XLSX)中未找到 '金额(元)' 列。")
            df_source['金额(元)'] = pd.to_numeric(df_source['金额(元)'].astype(str).str.replace('¥', '').str.strip(), errors='coerce')
            df_source['_source_tag'] = '#WechatPay'
        
        elif file_type == "AliPay":
            df_source = pd.read_csv(io_source, skiprows=skip_rows, encoding=csv_encoding)
            df_source.columns = df_source.columns.str.strip()
            amount_col_name = next((col for col in df_source.columns if "金额" in col), None)
            if not amount_col_name:
                raise ValueError("在支付宝文件中找不到任何 '金额' 列。")
            df_source.rename(columns={amount_col_name: '金额(元)'}, inplace=True)
            df_source.rename(columns=ALIPAY_MAPPING, inplace=True) 
            print(f"  ...已应用支付宝列名映射。")
            df_source['金额(元)'] = pd.to_numeric(df_source['金额(元)'], errors='coerce')
            df_source['_source_tag'] = '#AliPay'

        elif file_type == "Merged":
            df_source = pd.read_excel(io_source, skiprows=skip_rows, engine='openpyxl')
            df_source.columns = df_source.columns.str.strip()
            col_name = '金额(元)' if '金额(元)' in df_source.columns else '金额'
            if col_name not in df_source.columns:
                 raise ValueError("合并后账单(XLSX)中未找到 '金额(元)' 或 '金额' 列。")
            df_source['金额(元)'] = pd.to_numeric(df_source[col_name].astype(str).str.replace('¥', '').str.strip(), errors='coerce')
            df_source['_source_tag'] = ''

        if excel_workbook_object:
            excel_workbook_object.close()

        print(f"  成功读取 {len(df_source)} 行数据。", flush=True)
        return df_source
        
    except Exception as e:
        print(f"  加载数据时出错: {e}", flush=True)
        if excel_workbook_object: excel_workbook_object.close()
        return None

def process_descriptions(df: pd.DataFrame) -> pd.DataFrame:
    print("...[描述处理] 1/3: 设置默认值 (来自规则或商品)...", flush=True)
    if '描述' not in df.columns:
        df['描述'] = pd.NA 
    
    if '商品' not in df.columns:
        df['商品'] = pd.NA
        
    df['描述'] = df['描述'].fillna(df['商品'])

    print("...[描述处理] 2/3: 应用充电桩描述逻辑...", flush=True)
    is_charging = df['商品'].astype(str).str.contains("自助服务-充电桩", na=False)
    charging_desc_split = df.loc[is_charging, '商品'].str.split('/', n=1, expand=True)
    if not charging_desc_split.empty:
        valid_desc = charging_desc_split[1].str.strip()
        df.loc[is_charging, '描述'] = valid_desc.where(valid_desc.notna() & (valid_desc != ''), df.loc[is_charging, '商品'])
    
    print("...[描述处理] 3/3: 应用最终清理规则 (支付, 订单号, / , ...)...", flush=True)
    pattern_desc_clear = r"(?:支付|商户|二维码收款|收款|美团|编号|转账|错误|订单编号|\d{10,})"
    condition_desc = df['描述'].astype(str).str.contains(pattern_desc_clear, na=False, regex=True)
    
    pattern_memo_clear = r'(?:\d{10,}|订单编号)' 
    condition_memo = pd.Series(False, index=df.index) 
    if '备注' in df.columns:
        condition_memo = df['备注'].astype(str).str.contains(pattern_memo_clear, na=False, regex=True)
            
    rows_to_clear = condition_desc | condition_memo
    df.loc[rows_to_clear, '描述'] = ""
    
    print("...描述处理完成。", flush=True)
    return df

# ******************************************************
# *** 'process_transactions' 核心处理模块 ***
# ******************************************************

def _process_hardcoded_transfers(df_hardcoded_raw):
    print("\n--- 正在处理 (硬编码) 转账/提现 记录...", flush=True)
    list_of_transfer_dfs = []
    
    # --- Logic 1: "肖恩" 转账 ---
    df_sean_raw = df_hardcoded_raw[df_hardcoded_raw["交易对方"] == "肖恩"].copy()
    
    # 1A. 处理 "收入" (来自肖恩)
    df_income_sean = df_sean_raw[df_sean_raw["收/支"] == "收入"].copy()
    if not df_income_sean.empty:
        print(f"  检测到 {len(df_income_sean)} 条来自 '肖恩' 的收入...", flush=True)
        df_transfer_out = pd.DataFrame(index=df_income_sean.index)
        df_transfer_out['金额'] = pd.to_numeric(df_income_sean['金额(元)'], errors='coerce') * -1
        df_transfer_out['记录类型'] = "转出"
        df_transfer_out['账户'] = "零钱2"
        df_transfer_in = pd.DataFrame(index=df_income_sean.index)
        df_transfer_in['金额'] = pd.to_numeric(df_income_sean['金额(元)'], errors='coerce')
        df_transfer_in['记录类型'] = "转入"
        df_transfer_in['账户'] = "零钱3"
        df_transfers = pd.concat([df_transfer_out, df_transfer_in]) # (关键) 转出在前
        df_transfers = df_transfers.join(df_income_sean[['交易时间']])
        list_of_transfer_dfs.append(df_transfers)

    # 1B. 处理 "支出" (发往肖恩)
    df_expense_sean = df_sean_raw[df_sean_raw["收/支"] == "支出"].copy()
    if not df_expense_sean.empty:
        print(f"  检测到 {len(df_expense_sean)} 条发往 '肖恩' 的支出...", flush=True)
        df_transfer_out = pd.DataFrame(index=df_expense_sean.index)
        df_transfer_out['金额'] = pd.to_numeric(df_expense_sean['金额(元)'], errors='coerce') * -1
        df_transfer_out['记录类型'] = "转出"
        df_transfer_out['账户'] = "零钱3"
        df_transfer_in = pd.DataFrame(index=df_expense_sean.index)
        df_transfer_in['金额'] = pd.to_numeric(df_expense_sean['金额(元)'], errors='coerce')
        df_transfer_in['记录类型'] = "转入"
        df_transfer_in['账户'] = "零钱2"
        df_transfers = pd.concat([df_transfer_out, df_transfer_in]) # (关键) 转出在前
        df_transfers = df_transfers.join(df_expense_sean[['交易时间']])
        list_of_transfer_dfs.append(df_transfers)

    # --- Logic 2: "工商银行(9579)" 提现 ---
    df_icbc_raw = df_hardcoded_raw[df_hardcoded_raw["交易对方"] == "工商银行(9579)"].copy()
    if not df_icbc_raw.empty:
        print(f"  检测到 {len(df_icbc_raw)} 条 '工商银行(9579)' 提现记录...", flush=True)
        df_transfer_out = pd.DataFrame(index=df_icbc_raw.index)
        df_transfer_out['金额'] = pd.to_numeric(df_icbc_raw['金额(元)'], errors='coerce') * -1
        df_transfer_out['记录类型'] = "转出"
        df_transfer_out['账户'] = "零钱3"
        df_transfer_out['子类别'] = "提现"
        df_transfer_in = pd.DataFrame(index=df_icbc_raw.index)
        df_transfer_in['金额'] = pd.to_numeric(df_icbc_raw['金额(元)'], errors='coerce')
        df_transfer_in['记录类型'] = "转入"
        df_transfer_in['账户'] = "工商银行"
        df_transfer_in['子类别'] = "提现"
        df_transfers = pd.concat([df_transfer_out, df_transfer_in]) # (关键) 转出在前
        df_transfers = df_transfers.join(df_icbc_raw[['交易时间']])
        list_of_transfer_dfs.append(df_transfers)

    if not list_of_transfer_dfs:
        return pd.DataFrame()

    df_all_transfers = pd.concat(list_of_transfer_dfs)
    df_all_transfers['主类别'] = "转账"
    df_all_transfers['子类别'] = df_all_transfers['子类别'].fillna("转账")
    df_all_transfers['对象'] = ""
    df_all_transfers['项目'] = ""
    df_all_transfers['名称'] = ""
    df_all_transfers['描述'] = "" 
    df_all_transfers['商家'] = "" 
    df_all_transfers['标签'] = "" 
    df_all_transfers['日期'] = df_all_transfers['交易时间'].dt.strftime('%Y/%m/%d')
    df_all_transfers['时间'] = df_all_transfers['交易时间'].dt.strftime('%H:%M:%S')
    df_all_transfers['_Sort_Date'] = df_all_transfers['交易时间']
    df_all_transfers['币种'] = "CNY"
    df_all_transfers['手续费'] = 0
    df_all_transfers['折扣'] = 0
    
    return df_all_transfers

def _process_main_transactions(df_main_raw, df_rules):
    print(f"\n--- 正在处理 {len(df_main_raw)} 条 '主交易' (收入/支出/中性) 记录...", flush=True) 
    
    print("  ...正在应用筛选 1/2 (按 '当前状态')...", flush=True)
    status_condition = pd.Series(True, index=df_main_raw.index)
    if "当前状态" in df_main_raw.columns:
        excluded_statuses = ["已全额退款", "交易关闭"]
        status_condition = (~df_main_raw["当前状态"].isin(excluded_statuses))
    
    print("  ...正在应用筛选 2/2 (按 '收/支')...", flush=True)
    income_expense_condition = pd.Series(True, index=df_main_raw.index)
    if "收/支" in df_main_raw.columns:
        income_expense_condition = (df_main_raw["收/支"] != "不计收支")
        
    final_mask = status_condition & income_expense_condition
    df_expenses = df_main_raw[final_mask].copy()
    
    count_filtered = len(df_main_raw) - len(df_expenses)
    if count_filtered > 0:
         print(f"  ...已通过筛选 (状态/不计收支) 过滤掉 {count_filtered} 条记录。", flush=True)
         
    if df_expenses.empty:
        print("  (没有需要处理的主交易记录)")
        return pd.DataFrame()
    
    # 2B. 账户清理 (向量化)
    print("  正在进行性能优化 (向量化处理)...", flush=True)
    try:
        account_map_regex = {
            r'^(?:平安银行信用卡\(4946\)).*': '平安银行4946',
            r'^(?:工商银行储蓄卡\(9579\)).*': '工商银行',
            r'^零钱.*': '零钱3'
        }
        df_expenses['支付方式'] = df_expenses['支付方式'].astype(str).replace(account_map_regex, regex=True)
        print("  ...账户清理完成。", flush=True)
    except Exception as e:
        print(f"  !!! 账户清理步骤发生错误: {e} !!!", flush=True)

    # 2C. 规则匹配 (混合模式)
    print("  正在进行向量化规则匹配 (混合模式)...", flush=True)
    df_target = df_expenses.copy()
    
    # (!!!) 确保 '交易对方' 列存在
    if '交易对方' not in df_target.columns:
        df_target['交易对方'] = pd.NA
        
    df_target['商家(old)'] = df_target['交易对方'].astype(str).str.strip()
    
    df_rules_exact = df_rules[df_rules['is_regex'] == 0]
    df_rules_regex = df_rules[df_rules['is_regex'] == 1]
    
    print(f"  ... 1/2: 正在执行 {len(df_rules_exact)} 条精确匹配...", flush=True)
    cols_to_apply = [col for col in df_rules.columns if col not in ['商家(old)', 'is_regex']]
    rename_dict = {col: f'{col}_rule' for col in cols_to_apply}
    cols_to_rename_in_exact = {k: v for k, v in rename_dict.items() if k in df_rules_exact.columns}
    df_rules_exact_renamed = df_rules_exact.rename(columns=cols_to_rename_in_exact)
    df_target = pd.merge(
        df_target, 
        df_rules_exact_renamed.drop(columns='is_regex', errors='ignore'), 
        on='商家(old)', 
        how='left'
    )
    for col_renamed in rename_dict.values():
        if col_renamed not in df_target.columns:
            df_target[col_renamed] = pd.NA
    
    print("  ... 2/2: 正在为剩余行执行 {len(df_rules_regex)} 条正则匹配...", flush=True)
    
    primary_category_rule_col = None
    if '主类别_rule' in df_target.columns:
        primary_category_rule_col = '主类别_rule'
    elif '调整后-主类_rule' in df_target.columns:
        primary_category_rule_col = '调整后-主类_rule'
    
    unmatched_mask = pd.Series(True, index=df_target.index)
    if primary_category_rule_col:
        unmatched_mask = df_target[primary_category_rule_col].isna()
                     
    if unmatched_mask.any() and not df_rules_regex.empty:
        cols_to_apply_regex = [col for col in df_rules_regex.columns if col not in ['商家(old)', 'is_regex']]
        unmatched_data = df_target.loc[unmatched_mask, ['交易对方']]
        for _, rule in df_rules_regex.iterrows():
            pattern = rule['商家(old)']
            try:
                match_mask = unmatched_data['交易对方'].str.contains(pattern, na=False, regex=True)
                rows_to_update_idx = unmatched_data[match_mask].index
                if not rows_to_update_idx.empty:
                    for col in cols_to_apply_regex:
                        df_target.loc[rows_to_update_idx, f'{col}_rule'] = rule[col]
                    unmatched_data = unmatched_data.drop(rows_to_update_idx)
                    if unmatched_data.empty: break 
            except re.error as e:
                print(f"  !!! 正则规则错误: '{pattern}' -> {e} !!!", flush=True)
    
    print("  ...混合匹配完成。", flush=True)

    print("  ...正在从规则应用分类...", flush=True)
    df_target['主类别'] = df_target.get('调整后-主类_rule', pd.NA)
    df_target['子类别'] = df_target.get('调整后-子类_rule', pd.NA)
    df_target['商家'] = df_target.get('商家_rule', pd.NA) # (已修正)
    df_target['名称'] = df_target.get('名称_rule', pd.NA)
    df_target['描述'] = df_target.get('描述_rule', pd.NA)
    df_target['项目'] = df_target.get('项目_rule', pd.NA)
    df_target['标签'] = df_target.get('标签_rule', pd.NA)
    df_target['对象'] = df_target.get('对象_rule', pd.NA)
    df_target['记录类型'] = df_target.get('记录类型_rule', pd.NA)
    
    # 2D. 设置默认值 (现在这是安全的)
    print("  ...正在设置默认值 (记录类型, 商家)...", flush=True)
    df_target['记录类型'] = df_target['记录类型'].fillna(df_target['收/支'])
    df_target['商家'] = df_target['商家'].fillna(df_target['交易对方']) # 这是您的要求："首先按规则，否则按交易对方"

    # 2E. (已合并) 应用餐饮时间特殊逻辑
    print("  ...正在应用餐饮时间特殊逻辑...", flush=True)
    merchant_pattern = r"(?:三镇民生|兰州|丝路|沙县|永和四喜|水煎包|老乡鸡|混沌|馄饨|牛肉汤|热干面馆|黄蜀郎|鸡公煲|彩凤小吃店|烧烤|食堂|路边摊|麦香园|长沙臭豆腐)"
    merchant_mask = df_target['商家(old)'].astype(str).str.contains(merchant_pattern, na=False, regex=True) | \
                    df_target['商家'].astype(str).str.contains(merchant_pattern, na=False, regex=True)
    item_pattern = r"(?:东苑[一二]层|西区食堂)"
    item_mask = df_target['商品'].astype(str).str.contains(item_pattern, na=False, regex=True)
    condition_mask = (merchant_mask | item_mask) & (df_target['记录类型'] == '支出')
    
    if condition_mask.any():
        hour = df_target.loc[condition_mask, '交易时间'].dt.hour
        condlist = [
            (hour >= 6) & (hour < 10), (hour >= 11) & (hour < 16),
            (hour >= 16) & (hour < 21), (hour >= 21) | (hour < 2)
        ]
        choicelist_subcat = ["早餐", "午餐", "晚餐", "夜宵"]
        df_target.loc[condition_mask, 'temp_meal_subcat'] = np.select(condlist, choicelist_subcat, default=pd.NA)
        final_meal_mask = condition_mask & df_target['temp_meal_subcat'].notna()
        df_target.loc[final_meal_mask, '主类别'] = '饮食'
        df_target.loc[final_meal_mask, '项目'] = '食'
        df_target.loc[final_meal_mask, '子类别'] = df_target['temp_meal_subcat']
        df_target.drop(columns=['temp_meal_subcat'], inplace=True, errors='ignore')

    # (已合并) 应用充电桩特殊逻辑
    print("  ...正在应用充电桩特殊逻辑 (分类)...", flush=True)
    is_charging_flag = df_target['商品'].astype(str).str.contains("自助服务-充电桩", na=False)
    if is_charging_flag.any():
        df_target.loc[is_charging_flag, ['名称', '主类别', '子类别']] = ['充电', '交通', '加油充电']
    
    # 2F. 描述处理
    df_target = process_descriptions(df_target)

    # 2G. 格式化目标列
    print("  ...正在格式化目标列...", flush=True)
    df_target['金额'] = pd.to_numeric(df_target['金额(元)'], errors='coerce')
    is_expense_mask_final = (df_target['记录类型'] == '支出')
    df_target.loc[is_expense_mask_final, '金额'] = df_target.loc[is_expense_mask_final, '金额'] * -1
    df_target['日期'] = df_target['交易时间'].dt.strftime('%Y/%m/%d')
    df_target['时间'] = df_target['交易时间'].dt.strftime('%H:%M:%S')
    df_target['_Sort_Date'] = df_target['交易时间'] 
    
    print("  ...正在应用'收入'账户(/)到'零钱3'的特殊逻辑...", flush=True)
    special_logic_mask = (df_target['支付方式'] == "/") & (df_target['记录类型'] == "收入")
    df_target.loc[special_logic_mask, '支付方式'] = "零钱3"

    df_target['账户'] = df_target['支付方式'] # '支付方式' 列已被清理
    df_target['币种'] = "CNY"
    df_target['手续费'] = 0
    df_target['折扣'] = 0

    receivable_payable_mask = df_target['记录类型'].isin(['应收款项', '应付款项'])
    df_target.loc[receivable_payable_mask, '商家'] = ""
    
    return df_target

def process_transactions(df_source, df_rules, start_datetime_obj):
    # 3A. 初始清理和筛选
    
    if '交易时间' not in df_source.columns:
        print(f"警告：合并后的数据中缺少 '交易时间' 列。可能是 MAPPING 失败。将创建空列。", flush=True)
        df_source['交易时间'] = pd.NaT
        
    df_source['交易时间'] = df_source['交易时间'].apply(robust_date_converter)
    print("已使用智能转换函数处理 '交易时间' 列。", flush=True)

    if start_datetime_obj:
        df_source = df_source.loc[
            df_source['交易时间'].notna() & (df_source['交易时间'] >= start_datetime_obj)
        ].copy()
        print(f"应用时间筛选：剩余 {len(df_source)} 行。", flush=True)
    else:
        df_source = df_source.loc[df_source['交易时间'].notna()].copy()
        
    if df_source.empty:
        print("警告：没有符合时间条件的有效数据 (或支付宝日期映射失败)。", flush=True)
        return pd.DataFrame(columns=TARGET_HEADERS) 

    # 3B. "分诊"(Triage)：按业务逻辑分离
    
    if '交易对方' not in df_source.columns:
        df_source['交易对方'] = pd.NA
        
    sean_mask = (df_source["交易对方"] == "肖恩")
    icbc_mask = (df_source["交易对方"] == "工商银行(9579)")
    hardcoded_mask = sean_mask | icbc_mask 
    
    df_transfers_raw = df_source[hardcoded_mask].copy() 
    df_main_raw = df_source[~hardcoded_mask].copy()     
    
    list_of_final_dfs = []

    # 3C. 分别处理
    if not df_transfers_raw.empty:
        df_transfers_processed = _process_hardcoded_transfers(df_transfers_raw) 
        list_of_final_dfs.append(df_transfers_processed)
    else:
        print("\n--- 未检测到 '硬编码' (肖恩/提现) 记录 ---", flush=True) 

    if not df_main_raw.empty:
        df_main_processed = _process_main_transactions(df_main_raw, df_rules)
        list_of_final_dfs.append(df_main_processed)
    else:
        print("\n--- 未检测到 '主交易' (收入/支出/中性) 记录 ---", flush=True)
    
    # 3D. 合并和最终清理
    if not list_of_final_dfs:
        print("警告：没有可处理的'支出'或'转账'记录。", flush=True)
        return pd.DataFrame(columns=TARGET_HEADERS) 

    print("\n--- 正在合并所有已处理的记录...", flush=True)
    df_final = pd.concat(list_of_final_dfs, ignore_index=True)

    fill_na_cols = ['主类别', '子类别', '名称', '项目', '标签', '对象']
    for col in fill_na_cols:
        if col not in df_final.columns:
            df_final[col] = "" 
    df_final[fill_na_cols] = df_final[fill_na_cols].fillna("")
    
    print("...正在应用来源标签 (Wechat/AliPay)...", flush=True)
    if '_source_tag' not in df_final.columns:
        df_final['_source_tag'] = ""
    df_final['_source_tag'] = df_final['_source_tag'].fillna("")
    df_final['标签'] = df_final['标签'].fillna("") 

    is_taggable_mask = ~df_final['记录类型'].isin(['转入', '转出'])
    source_tags_to_apply = df_final['_source_tag'].where(is_taggable_mask, "")

    df_final['标签'] = df_final['标签'].str.cat(source_tags_to_apply, sep=' ').str.strip()
    df_final['标签'] = df_final['标签'].str.replace(r'\s+', ' ', regex=True)
    
    print("...向量化处理完成。", flush=True)
    
    final_cols = TARGET_HEADERS + ['_Sort_Date']
    df_final = df_final[[col for col in final_cols if col in df_final.columns]]
    
    return df_final


# --- 主函数 ---

def main():
    try:
        source_files, start_date = get_user_inputs()
        if not source_files:
            return

        if not TARGET_DIR.exists():
            try:
                TARGET_DIR.mkdir(parents=True, exist_ok=True)
                print(f"已自动创建目标文件夹: {TARGET_DIR}", flush=True)
            except Exception as e:
                print(f"创建文件夹失败: {e}", flush=True)
                return

        df_rules = load_rules(RULE_BOOK_PATH)
        if df_rules is None:
            return
            
        list_of_dfs = []
        print("\n--- 开始批量加载文件 ---", flush=True)
        for file_path in source_files:
            df_source = sniff_and_load_data(file_path)
            if df_source is not None:
                list_of_dfs.append(df_source)
            else:
                print(f"  {bcolors.WARNING}警告: 文件 {file_path.name} 加载失败，已跳过。{bcolors.ENDC}", flush=True)
        
        if not list_of_dfs:
            print(f"{bcolors.FAIL}错误：所有文件均加载失败，程序退出。{bcolors.ENDC}", flush=True)
            return
            
        print("\n--- 文件加载完毕，正在合并... ---", flush=True)
        df_all_sources = pd.concat(list_of_dfs, ignore_index=True)
        print(f"合并后总行数: {len(df_all_sources)} (若支付宝为0, 请检查CSV列名是否在MAPPING中)", flush=True)

        df_final = process_transactions(df_all_sources, df_rules, start_date)
        
        if df_final.empty:
            print("没有可导出的数据，程序已完成。", flush=True)
            return

        # (关键) 稳定排序逻辑
        print("正在创建稳定排序键 (用于转账)...", flush=True)
        sort_key_map = {'转出': 0, '转入': 1}
        df_final['_Transfer_Sort_Key'] = df_final['记录类型'].map(sort_key_map).fillna(9)

        print("正在按日期降序排列 (最新在上, 转出在前)...", flush=True)
        df_final.sort_values(
            by=["_Sort_Date", "_Transfer_Sort_Key"], 
            ascending=[False, True], 
            inplace=True,
            kind='stable'
        )
        
        if '_Transfer_Sort_Key' in df_final.columns:
             df_final.drop(columns=['_Transfer_Sort_Key'], inplace=True)
        
        if '_Sort_Date' in df_final.columns:
             df_final.drop(columns=['_Sort_Date'], inplace=True)
        
        df_final.reset_index(drop=True, inplace=True)

        # (已重新应用颜色)
        print(f"\n--- {bcolors.BOLD}最终文件检查 (基于最终CSV行号){bcolors.ENDC} ---", flush=True)
        
        uncategorized_mask = (df_final['主类别'] == "") | (df_final['子类别'] == "")
        uncategorized_count = uncategorized_mask.sum()
        if uncategorized_count > 0:
            print(f"  {bcolors.WARNING}[警告] 发现 {uncategorized_count} 条记录 '主类别' 或 '子类别' 为空:{bcolors.ENDC}", flush=True)
            problem_rows_uncat = df_final[uncategorized_mask].copy()
            problem_rows_uncat['print_amount'] = problem_rows_uncat['金额'].apply(lambda x: f"{x:.2f}")
            print(f"  {bcolors.WARNING}--- 未分类的记录 (最多显示前 10 条) ---{bcolors.ENDC}", flush=True)
            for row in problem_rows_uncat.head(10).itertuples():
                csv_row_num = row.Index + 2
                print(f"  {bcolors.WARNING}- [CSV 第 {csv_row_num} 行] [日期: {row.日期}, 商家: {row.商家}, 金额: {row.print_amount}]{bcolors.ENDC}", flush=True)
            if uncategorized_count > 10:
                print(f"  {bcolors.WARNING}... (及其他 {uncategorized_count - 10} 条){bcolors.ENDC}", flush=True)
        else:
            print(f"  {bcolors.OKGREEN}[通过] 所有记录均已分类 (主类别/子类别 均不为空)。{bcolors.ENDC}", flush=True)

        receivable_payable_mask = df_final['记录类型'].isin(['应收款项', '应付款项'])
        missing_object_mask = (receivable_payable_mask) & (df_final['对象'] == "")
        missing_object_count = missing_object_mask.sum()
        if missing_object_count > 0:
            print(f"  {bcolors.WARNING}[警告] 发现 {missing_object_count} 条 '应收款项'/'应付款项' 记录缺少 '对象':{bcolors.ENDC}", flush=True)
            problem_rows_obj = df_final[missing_object_mask].copy()
            problem_rows_obj['print_amount'] = problem_rows_obj['金额'].apply(lambda x: f"{x:.2f}")
            print(f"  {bcolors.WARNING}--- 缺少'对象'的记录 (最多显示前 10 条) ---{bcolors.ENDC}", flush=True)
            for row in problem_rows_obj.head(10).itertuples():
                csv_row_num = row.Index + 2
                print(f"  {bcolors.WARNING}- [CSV 第 {csv_row_num} 行] [日期: {row.日期}, 描述: {row.描述}, 金额: {row.print_amount}]{bcolors.ENDC}", flush=True)
            if missing_object_count > 10:
                print(f"  {bcolors.WARNING}... (及其他 {missing_object_count - 10} 条){bcolors.ENDC}", flush=True)
        else:
            print(f"  {bcolors.OKGREEN}[通过] 所有 '应收款项'/'应付款项' 记录均已填写 '对象'。{bcolors.ENDC}", flush=True)

        print("-" * 20, flush=True)
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        target_path = TARGET_DIR / f"MOZE导入_{timestamp}.csv"
        
        df_final.to_csv(
            target_path, 
            index=False, 
            columns=TARGET_HEADERS, 
            encoding='utf-8-sig'
        )
        
        print("-" * 30, flush=True)
        print(f"处理完成！已成功生成文件: '{target_path}'", flush=True)  
        print(f"总共写入了 {len(df_final)} 条记录。", flush=True)

    except Exception as e:
        print(f"\n{bcolors.FAIL}处理过程中发生未捕获的严重错误: {e}{bcolors.ENDC}", flush=True)
        print("请检查：", flush=True)
        print("1. 您选择的 Excel/CSV 文件是否是支持的格式？", flush=True)
        print(f"2. Z {RULE_BOOK_PATH.name} 格式是否正确 (是否添加了 'is_regex' 列)？", flush=True)
        print(f"3. 确认您有权限写入 '{TARGET_DIR}' 文件夹。", flush=True)
        print(f"\n{bcolors.FAIL}--- 详细错误信息 ---{bcolors.ENDC}")
        traceback.print_exc()
        print(f"{bcolors.FAIL}---------------------\n{bcolors.ENDC}")


if __name__ == "__main__":
    main()
    input("\n...操作已完成，按 [Enter] 键退出。")