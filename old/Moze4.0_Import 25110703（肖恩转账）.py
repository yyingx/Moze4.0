import pandas as pd
import os
import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog
import openpyxl  # 用于 "嗅探" 和读取 Excel
import csv       # 用于 "嗅探" CSV
from pathlib import Path  # 导入现代路径处理库
import traceback # 用于打印详细错误
import re # 导入正则表达式库
import numpy as np # 导入 numpy

# --- ANSI 颜色代码 ---
class bcolors:
    OKGREEN = '\033[92m' # 绿色 (通过)
    WARNING = '\033[91m' # 红色 (警告)
    FAIL = '\033[91m'    # 红色 (失败)
    ENDC = '\033[0m'     # 结束符
    BOLD = '\033[1m'     # 加粗

# --- 常量定义 ---
RULE_BOOK_PATH = Path(r"E:\天之逸2025\Moze4.0\Moze Dict.xlsx")
TARGET_DIR = Path(r"E:\天之逸2025\Moze4.0\Moze4.0_Import")


# 支付宝列名映射
ALIPAY_MAPPING = {
    '交易时间': '交易时间',
    '交易分类': '交易类型',
    '交易对方': '交易对方',
    '商品说明': '商品',
    '收/支': '收/支',
    '收/付款方式': '支付方式',
    '交易状态': '当前状态',
    '交易订单号': '交易单号',
    '商家订单号': '商户单号',
    '备注': '备注'
}

# 目标 CSV 抬头
TARGET_HEADERS = [
    "账户", "币种", "记录类型", "主类别", "子类别", "金额",  
    "手续费", "折扣", "名称", "商家", "日期", "时间",  
    "项目", "描述", "标签", "对象"
]

# --- 辅助函数 ---

def robust_date_converter(x):
    """
    智能处理混合类型的日期列 (字符串, Excel浮点数, 或 datetime)
    """
    if isinstance(x, (datetime.datetime, datetime.date)):
        return pd.to_datetime(x)
    elif isinstance(x, (int, float)):
        try:
            return pd.to_datetime(x, unit='D', origin='1899-12-30')
        except Exception:
            return pd.NaT
    elif isinstance(x, str):
        return pd.to_datetime(x, errors='coerce')
    else:
        return pd.NaT

def get_user_inputs():
    """
    使用 Tkinter 弹出对话框，获取(多个)源文件路径和起始日期。
    返回: (list[Path], datetime) 或 (None, None)
    """
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
        "**格式： yymmdd**\n"
        "（例如: 251001）\n\n"
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
    """
    从指定路径加载规则文件 (Moze Dict.xlsx)，并处理 'is_regex' 列
    返回: df_rules 或 None
    """
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
            print("提示：规则文件中未找到 'is_regex' 列，将全部视为精确匹配。", flush=True)
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
    """
    分析文件类型 (嗅探), 加载数据, 并统一列名。
    (此版本会为 微信/支付宝 添加 _source_tag)
    返回: df_source 或 None
    """
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
                        file_type, skip_rows, csv_encoding = "Alipay", i + 1, 'utf-8'
                        print(f"  检测到：【原始支付宝文件 (UTF-8)】, 抬头在第 {i+1} 行。", flush=True)
                        break
        except UnicodeDecodeError:
            try:
                with open(file_path, 'r', encoding='GBK', errors='ignore') as f:
                    for i, line in enumerate(f):
                        if i > 30: break
                        if keyword in line:
                            file_type, skip_rows, csv_encoding = "Alipay", i + 1, 'GBK'
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
            elif "微信支付账单明细" in cell_A1:
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
            
            df_source['金额(元)'] = pd.to_numeric(
                df_source['金额(元)'].astype(str).str.replace('¥', '').str.strip(),
                errors='coerce'
            )
            df_source['_source_tag'] = '#WechatPay' # 添加来源标签
        
        elif file_type == "Alipay":
            df_source = pd.read_csv(io_source, skiprows=skip_rows, encoding=csv_encoding)
            df_source.columns = df_source.columns.str.strip()
            
            amount_col_name = next((col for col in df_source.columns if "金额" in col), None)
            if not amount_col_name:
                raise ValueError("在支付宝文件中找不到任何 '金额' 列。")
                
            df_source.rename(columns={amount_col_name: '金额(元)'}, inplace=True)
            df_source.rename(columns=ALIPAY_MAPPING, inplace=True)
            df_source['金额(元)'] = pd.to_numeric(df_source['金额(元)'], errors='coerce')
            df_source['_source_tag'] = '#Alipay' # 添加来源标签

        elif file_type == "Merged":
            df_source = pd.read_excel(io_source, skiprows=skip_rows, engine='openpyxl')
            df_source.columns = df_source.columns.str.strip()
            
            if '金额(元)' not in df_source.columns:
                 if '金额' in df_source.columns:
                     col_name = '金额'
                 else:
                     raise ValueError("合并后账单(XLSX)中未找到 '金额(元)' 或 '金额' 列。")
            else:
                col_name = '金额(元)'
                 
            df_source['金额(元)'] = pd.to_numeric(
                df_source[col_name].astype(str).str.replace('¥', '').str.strip(),
                errors='coerce'
            )
            df_source['_source_tag'] = '' # 合并文件没有特定来源

        if excel_workbook_object:
            excel_workbook_object.close()
            print("  ... (已关闭 Excel 文件句柄)", flush=True)

        print(f"  成功读取 {len(df_source)} 行数据。", flush=True)
        return df_source
        
    except Exception as e:
        print(f"  加载数据时出错: {e}", flush=True)
        if excel_workbook_object: excel_workbook_object.close()
        return None

def process_descriptions(df: pd.DataFrame) -> pd.DataFrame:
    """
    按顺序应用所有“描述”列的逻辑
    """
    
    print("...[描述处理] 1/3: 设置默认值 (来自规则或商品)...", flush=True)
    df['描述'] = df['描述'].fillna(df['商品'])

    print("...[描述处理] 2/3: 应用充电桩描述逻辑...", flush=True)
    is_charging = df['商品'].str.contains("自助服务-充电桩", na=False)
    
    charging_desc_split = df.loc[is_charging, '商品'].str.split('/', n=1, expand=True)
    if not charging_desc_split.empty:
        valid_desc = charging_desc_split[1].str.strip()
        df.loc[is_charging, '描述'] = valid_desc.where(
            valid_desc.notna() & (valid_desc != ''), 
            df.loc[is_charging, '商品'] # Fallback
        )
    
    print("...[描述处理] 3/3: 应用最终清理规则 (支付, 订单号, / , ...)...", flush=True)
    
    pattern_desc_clear = r"支付|商户|二维码收款|收款|美团|编号|转账|/|错误|订单编号|\d{10,}"
    condition_desc = df['描述'].astype(str).str.contains(
        pattern_desc_clear, 
        na=False, 
        regex=True
    )
    
    pattern_memo_clear = r'\d{10,}|订单编号' 
    condition_memo = pd.Series(False, index=df.index) 
    
    if '备注' in df.columns:
        condition_memo = df['备注'].astype(str).str.contains(
            pattern_memo_clear,
            na=False, 
            regex=True
        )
            
    rows_to_clear = condition_desc | condition_memo
    df.loc[rows_to_clear, '描述'] = ""
    
    print("...描述处理完成。", flush=True)
    return df

# ******************************************************
# *** (已重构) 'process_transactions' 函数 ***
# *** (此版本分离了 "支出" 和 "转账" 逻辑) ***
# ******************************************************
def process_transactions(df_source, df_rules, start_datetime_obj):
    """
    核心处理函数：清理数据、筛选、并使用向量化应用规则 (混合模式)。
    """
    
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
        print("警告：没有符合时间条件的有效数据。", flush=True)
        return pd.DataFrame(columns=TARGET_HEADERS) 

    list_of_final_dfs = []
    
    df_expenses_raw = df_source[df_source["收/支"] == "支出"].copy()
    df_income_raw = df_source[df_source["收/支"] == "收入"].copy()

    # ******************************************************
    # *** 块 1: 处理 "收入" (来自 肖恩) ***
    # ******************************************************
    if not df_income_raw.empty:
        print("\n--- 正在处理 '收入' 记录 (用于转账)...", flush=True)
        transfer_mask = (df_income_raw["交易对方"] == "肖恩")
        df_income_sean = df_income_raw[transfer_mask].copy()
        
        if not df_income_sean.empty:
            print(f"  检测到 {len(df_income_sean)} 条来自 '肖恩' 的收入，将(硬编码)处理为转账...", flush=True)

            # A. 创建 "转出" (Transfer Out) 行 (按截图: 负数)
            df_transfer_out = pd.DataFrame(index=df_income_sean.index)
            df_transfer_out['金额'] = pd.to_numeric(df_income_sean['金额(元)'], errors='coerce') * -1 # 变为负数
            df_transfer_out['记录类型'] = "转出"
            df_transfer_out['主类别'] = "转账"
            df_transfer_out['子类别'] = "转账"
            df_transfer_out['账户'] = "零钱2"
            df_transfer_out['对象'] = ""

            # B. 创建 "转入" (Transfer In) 行 (按截图: 正数)
            df_transfer_in = pd.DataFrame(index=df_income_sean.index)
            df_transfer_in['金额'] = pd.to_numeric(df_income_sean['金额(元)'], errors='coerce') # 保持正数
            df_transfer_in['记录类型'] = "转入"
            df_transfer_in['主类别'] = "转账"
            df_transfer_in['子类别'] = "转账"
            df_transfer_in['账户'] = "零钱3"
            df_transfer_in['对象'] = "" 
            
            df_transfers = pd.concat([df_transfer_out, df_transfer_in])
            
            df_transfers = df_transfers.join(df_income_sean[['交易时间']])
            
            df_transfers['项目'] = ""
            df_transfers['名称'] = ""
            df_transfers['描述'] = "" 
            df_transfers['商家'] = "" 
            df_transfers['标签'] = "" 
            
            df_transfers['日期'] = df_transfers['交易时间'].dt.strftime('%Y/%m/%d')
            df_transfers['时间'] = df_transfers['交易时间'].dt.strftime('%H:%M:%S')
            df_transfers['_Sort_Date'] = df_transfers['交易时间'] 
            df_transfers['币种'] = "CNY"
            df_transfers['手续费'] = 0
            df_transfers['折扣'] = 0
            
            list_of_final_dfs.append(df_transfers)

    # ******************************************************
    # *** 块 2: 处理 "支出" (所有现有逻辑) ***
    # ******************************************************
    if not df_expenses_raw.empty:
        print("\n--- 正在处理 '支出' 记录...", flush=True)
        status_condition = pd.Series(True, index=df_expenses_raw.index)
        if "当前状态" in df_expenses_raw.columns:
            excluded_statuses = ["已全额退款", "交易关闭"]
            status_condition = (~df_expenses_raw["当前状态"].isin(excluded_statuses))
            
        df_expenses = df_expenses_raw[status_condition].copy()
        
        # (新增) 分离 "支出 -> 肖恩" 的转账
        expense_transfer_mask = (df_expenses["交易对方"] == "肖恩")
        df_expense_sean = df_expenses[expense_transfer_mask].copy()
        df_expenses_others = df_expenses[~expense_transfer_mask].copy() # 剩余的都是真实支出
        
        # ******************************************************
        # *** (新增) 块 2A: 处理 "支出 -> 肖恩" (转账) ***
        # ******************************************************
        if not df_expense_sean.empty:
            print(f"  检测到 {len(df_expense_sean)} 条发往 '肖恩' 的支出，将(硬编码)处理为转账...", flush=True)

            # A. 创建 "转出" (Transfer Out) 行 (按要求: 负数)
            df_transfer_out = pd.DataFrame(index=df_expense_sean.index)
            df_transfer_out['金额'] = pd.to_numeric(df_expense_sean['金额(元)'], errors='coerce') * -1 # 保持负数
            df_transfer_out['记录类型'] = "转出"
            df_transfer_out['主类别'] = "转账"
            df_transfer_out['子类别'] = "转账"
            df_transfer_out['账户'] = "零钱3" # (新)
            df_transfer_out['对象'] = "" 

            # B. 创建 "转入" (Transfer In) 行 (按要求: 正数)
            df_transfer_in = pd.DataFrame(index=df_expense_sean.index)
            df_transfer_in['金额'] = pd.to_numeric(df_expense_sean['金额(元)'], errors='coerce') # 变为正数
            df_transfer_in['记录类型'] = "转入"
            df_transfer_in['主类别'] = "转账"
            df_transfer_in['子类别'] = "转账"
            df_transfer_in['账户'] = "零钱2" # (新)
            df_transfer_in['对象'] = "" 
            
            df_transfers = pd.concat([df_transfer_out, df_transfer_in])
            
            df_transfers = df_transfers.join(df_expense_sean[['交易时间']])
            
            df_transfers['项目'] = ""
            df_transfers['名称'] = ""
            df_transfers['描述'] = "" 
            df_transfers['商家'] = "" 
            df_transfers['标签'] = "" 
            
            df_transfers['日期'] = df_transfers['交易时间'].dt.strftime('%Y/%m/%d')
            df_transfers['时间'] = df_transfers['交易时间'].dt.strftime('%H:%M:%S')
            df_transfers['_Sort_Date'] = df_transfers['交易时间'] 
            df_transfers['币种'] = "CNY"
            df_transfers['手续费'] = 0
            df_transfers['折扣'] = 0
            
            list_of_final_dfs.append(df_transfers)

        # ******************************************************
        # *** 块 2B: 处理 "其他支出" (所有现有逻辑) ***
        # ******************************************************
        if not df_expenses_others.empty:
            print(f"  已筛选出 {len(df_expenses_others)} 条 '其他支出' 记录进行处理...", flush=True)
            
            df_expenses = df_expenses_others # 重新赋值，供后续逻辑使用

            print("  正在进行性能优化 (向量化处理)...", flush=True)
            try:
                account_map_regex = {
                    r'^平安银行信用卡\(4946\).*': '平安银行4946',
                    r'^工商银行储蓄卡\(9579\).*': '工商银行',
                    r'^零钱.*': '零钱3'
                    
                }
                df_expenses['支付方式'] = df_expenses['支付方式'].astype(str).replace(account_map_regex, regex=True)
                print("  ...账户清理完成。", flush=True)
            except Exception as e:
                print(f"  !!! {bcolors.FAIL}账户清理步骤发生错误: {e}{bcolors.ENDC} !!!", flush=True)

            print("  正在进行向量化规则匹配 (混合模式)...", flush=True)
            df_target = df_expenses.copy()
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
            
            print(f"  ... 2/2: 正在为剩余行执行 {len(df_rules_regex)} 条正则匹配...", flush=True)
            unmatched_mask = df_target['主类别_rule'].isna()
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
                        print(f"  !!! {bcolors.FAIL}正则规则错误: '{pattern}' -> {e}{bcolors.ENDC} !!!", flush=True)
            
            print("  ...混合匹配完成。", flush=True)

            df_target.rename(columns={
                '名称_rule': '名称', '商家_rule': '商家', '描述_rule': '描述',
                '主类别_rule': '主类别', '子类别_rule': '子类别', '项目_rule': '项目',
                '标签_rule': '标签', '对象_rule': '对象', '记录类型_rule': '记录类型'
            }, inplace=True)
            
            df_target['记录类型'] = df_target['记录类型'].fillna("支出")
            df_target['商家'] = df_target['商家'].fillna(df_target['交易对方'])

            print("  ...正在应用餐饮时间特殊逻辑...", flush=True)
            merchant_pattern = r"三镇民生|兰州|丝路|沙县|永和四喜|水煎包|老乡鸡|牛肉汤|热干面馆|黄蜀郎|鸡公煲|彩凤小吃店|烧烤|食堂|路边摊|麦香园|长沙臭豆腐"
            merchant_mask_old = df_target['商家(old)'].str.contains(merchant_pattern, na=False, regex=True)
            merchant_mask_new = df_target['商家'].str.contains(merchant_pattern, na=False, regex=True)
            merchant_mask = merchant_mask_old | merchant_mask_new
            item_pattern = r"东苑一层|西区食堂"
            item_mask = df_target['商品'].str.contains(item_pattern, na=False, regex=True)
            condition_mask = merchant_mask | item_mask
            hour = df_target['交易时间'].dt.hour
            is_expense_mask = (df_target['记录类型'] == '支出')
            condlist = [
                (hour >= 6) & (hour < 10), (hour >= 11) & (hour < 16),
                (hour >= 16) & (hour < 21), (hour >= 21) | (hour < 2)
            ]
            choicelist_subcat = ["早餐", "午餐", "晚餐", "夜宵"]
            df_target['temp_meal_subcat'] = np.select(condlist, choicelist_subcat, default=pd.NA)
            final_meal_mask = (condition_mask & is_expense_mask & df_target['temp_meal_subcat'].notna())
            df_target.loc[final_meal_mask, '主类别'] = '饮食'
            df_target.loc[final_meal_mask, '项目'] = '食'
            df_target.loc[final_meal_mask, '子类别'] = df_target['temp_meal_subcat']
            df_target.drop(columns=['temp_meal_subcat'], inplace=True)
            
            print("  ...正在应用充电桩特殊逻辑 (分类)...", flush=True)
            is_charging_flag = df_target['商品'].str.contains("自助服务-充电桩", na=False)
            df_target.loc[is_charging_flag, ['名称', '主类别', '子类别']] = ['充电', '交通', '加油充电']
            
            df_target = process_descriptions(df_target)

            print("  ...正在格式化目标列...", flush=True)
            df_target['金额'] = pd.to_numeric(df_target['金额(元)'], errors='coerce') * -1
            df_target['日期'] = df_target['交易时间'].dt.strftime('%Y/%m/%d')
            df_target['时间'] = df_target['交易时间'].dt.strftime('%H:%M:%S')
            df_target['_Sort_Date'] = df_target['交易时间'] 
            df_target['账户'] = df_target['支付方式'] 
            df_target['币种'] = "CNY"
            df_target['手续费'] = 0
            df_target['折扣'] = 0

            print("  ...正在应用'应收/应付'商户特殊规则...", flush=True)
            receivable_payable_mask = df_target['记录类型'].isin(['应收款项', '应付款项'])
            df_target.loc[receivable_payable_mask, '商家'] = ""
            
            list_of_final_dfs.append(df_target)
    
    # ******************************************************
    # *** 块 3: 合并和最终清理 ***
    # ******************************************************
    if not list_of_final_dfs:
        print("警告：没有可处理的'支出'或'转账'记录。", flush=True)
        return pd.DataFrame(columns=TARGET_HEADERS) 

    print("\n--- 正在合并 支出 和 转账 记录...", flush=True)
    df_final = pd.concat(list_of_final_dfs, ignore_index=True)

    fill_na_cols = ['主类别', '子类别', '名称', '项目', '标签', '对象']
    for col in fill_na_cols:
        if col not in df_final.columns:
            df_final[col] = "" 
    df_final[fill_na_cols] = df_final[fill_na_cols].fillna("")
    
    print("...正在应用来源标签 (Wechat/Alipay)...", flush=True)
    if '_source_tag' not in df_final.columns:
        df_final['_source_tag'] = ""
    df_final['_source_tag'] = df_final['_source_tag'].fillna("")
    df_final['标签'] = df_final['标签'].fillna("") 

    # (修复) 转账记录 (转入/转出) 不应被自动打上来源标签
    is_taggable_mask = df_final['记录类型'].isin(['支出'])
    source_tags_to_apply = df_final['_source_tag'].where(is_taggable_mask, "")

    df_final['标签'] = df_final['标签'].str.cat(source_tags_to_apply, sep=' ').str.strip()
    df_final['标签'] = df_final['标签'].str.replace(r'\s+', ' ', regex=True)
    
    print("...向量化处理完成。", flush=True)
    
    final_cols = TARGET_HEADERS + ['_Sort_Date']
    df_final = df_final[[col for col in final_cols if col in df_final.columns]]
    
    return df_final


# --- 主函数 ---

def main():
    """
    主程序执行流程
    """
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
        print(f"合并后总行数: {len(df_all_sources)}", flush=True)

        df_final = process_transactions(df_all_sources, df_rules, start_date)
        
        if df_final.empty:
            print("没有可导出的数据，程序已完成。", flush=True)
            return

        print("正在按日期降序排列 (最新在上)...", flush=True)
        df_final.sort_values(by="_Sort_Date", ascending=False, inplace=True)
        
        df_final.reset_index(drop=True, inplace=True)

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
                print(f"  {bcolors.WARNING}  - [CSV 第 {csv_row_num} 行] [日期: {row.日期}, 商家: {row.商家}, 金额: {row.print_amount}]{bcolors.ENDC}", flush=True)
            
            if uncategorized_count > 10:
                print(f"  {bcolors.WARNING}  ... (及其他 {uncategorized_count - 10} 条){bcolors.ENDC}", flush=True)

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
                print(f"  {bcolors.WARNING}  - [CSV 第 {csv_row_num} 行] [日期: {row.日期}, 描述: {row.描述}, 金额: {row.print_amount}]{bcolors.ENDC}", flush=True)
            
            if missing_object_count > 10:
                print(f"  {bcolors.WARNING}  ... (及其他 {missing_object_count - 10} 条){bcolors.ENDC}", flush=True)

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
        print(f"2. 您的 {RULE_BOOK_PATH.name} 格式是否正确 (是否添加了 'is_regex' 列)？", flush=True)
        print(f"3. 确认您有权限写入 '{TARGET_DIR}' 文件夹。", flush=True)
        print(f"\n{bcolors.FAIL}--- 详细错误信息 ---{bcolors.ENDC}")
        traceback.print_exc()
        print(f"{bcolors.FAIL}---------------------\n{bcolors.ENDC}")


if __name__ == "__main__":
    main()
    input("\n...操作已完成，按 [Enter] 键退出。")