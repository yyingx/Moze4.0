import pandas as pd
import os
import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog
import openpyxl  # 用于 "嗅探" 和读取 Excel
import csv       # 用于 "嗅探" CSV
from pathlib import Path  # 导入现代路径处理库
import traceback # 用于打印详细错误

# --- * 1. 常量定义 * ---
# (请根据您的实际路径修改)
RULE_BOOK_PATH = Path(r"D:\OneDrive\桌面\Moze4.0\Moze Dict.csv")
TARGET_DIR = Path(r"D:\OneDrive\桌面\Moze4.0\Moze4.0_Import")

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

# --- * 2. 辅助函数 * ---

def robust_date_converter(x):
    """
    智能处理混合类型的日期列 (字符串, Excel浮点数, 或 datetime)
    """
    if isinstance(x, (datetime.datetime, datetime.date)):
        return pd.to_datetime(x)
    elif isinstance(x, (int, float)):
        try:
            # Excel 日期基准
            return pd.to_datetime(x, unit='D', origin='1899-12-30')
        except Exception:
            return pd.NaT
    elif isinstance(x, str):
        return pd.to_datetime(x, errors='coerce')
    else:
        return pd.NaT

def get_user_inputs():
    """
    使用 Tkinter 弹出对话框，获取源文件路径和起始日期。
    返回: (Path, datetime) 或 (None, None)
    """
    root = tk.Tk()
    root.withdraw() # 隐藏主窗口

    print("第1步: 请选择您的账单文件 (微信/支付宝/合并后 均可)...", flush=True)
    source_file_str = filedialog.askopenfilename(
        title="第1步: 请选择您的账单文件 (微信/支付宝/合并后 均可)",
        filetypes=(("Excel/CSV 文件", "*.xlsx;*.csv"), ("所有文件", "*.*"))
    )
    
    if not source_file_str:
        print("操作已取消。", flush=True)
        root.destroy()
        return None, None
    
    print(f"您选择的源文件是: {source_file_str}", flush=True)
    
    start_time_str = simpledialog.askstring(
        "第2步: 筛选条件 (按日期)",
        "请输入起始日期（包含当天及之后）\n\n"
        "**格式： yymmdd**\n"
        "（例如: 251001）\n\n"
        "*** 如果留空，将导入所有时间的数据 ***",
        parent=root
    )
    root.destroy() # 关闭所有对话框

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

    return Path(source_file_str), start_datetime_obj

def load_rules(rule_path: Path):
    """
    从指定路径加载规则文件 (Moze Dict.csv)
    返回: df_rules 或 None
    """
    if not rule_path.exists():
        print(f"错误：找不到规则文件！ {rule_path}", flush=True)
        return None
        
    print(f"正在读取规则文件: {rule_path}", flush=True)
    try:
        df_rules = pd.read_csv(rule_path, encoding='utf-8-sig')
        
        if '商家(old)' not in df_rules.columns:
            print(f"错误：规则文件 {rule_path} 中缺少 '商家(old)' 列！", flush=True)
            return None
            
        df_rules['商家(old)'] = df_rules['商家(old)'].str.strip()
        df_rules.drop_duplicates(subset=['商家(old)'], keep='last', inplace=True)
        
        print(f"成功加载 {len(df_rules)} 条[唯一]分类规则。", flush=True)
        return df_rules
    except Exception as e:
        print(f"读取规则文件时出错: {e}", flush=True)
        return None

def sniff_and_load_data(file_path: Path):
    """
    分析文件类型 (嗅探), 加载数据, 并统一列名。
    返回: df_source 或 None
    """
    print(f"正在分析并加载文件: {file_path}", flush=True)
    file_type = None
    skip_rows = 0
    df_source = None
    excel_workbook_object = None
    csv_encoding = None

    if file_path.suffix == '.csv':
        # --- 嗅探CSV (支付宝) ---
        keyword = "支付宝支付科技有限公司"
        try:
            # 1. 尝试 UTF-8
            with open(file_path, 'r', encoding='utf-8') as f:
                for i, line in enumerate(f):
                    if i > 30: break
                    if keyword in line:
                        file_type, skip_rows, csv_encoding = "Alipay", i + 1, 'utf-8'
                        print(f"检测到：【原始支付宝文件 (UTF-8)】, 抬头在第 {i+1} 行。", flush=True)
                        break
        except UnicodeDecodeError:
            # 2. 尝试 GBK
            try:
                with open(file_path, 'r', encoding='GBK', errors='ignore') as f:
                    for i, line in enumerate(f):
                        if i > 30: break
                        if keyword in line:
                            file_type, skip_rows, csv_encoding = "Alipay", i + 1, 'GBK'
                            print(f"检测到：【原始支付宝文件 (GBK)】, 抬头在第 {i+1} 行。", flush=True)
                            break
            except Exception as e:
                print(f"GBK 编码也失败: {e}", flush=True)

    elif file_path.suffix == '.xlsx':
        # --- 嗅探XLSX (微信 / 合并后) ---
        try:
            excel_workbook_object = openpyxl.load_workbook(file_path, read_only=True)
            sheet = excel_workbook_object.active
            cell_A1 = str(sheet['A1'].value)
            
            if cell_A1 == "交易时间":
                file_type, skip_rows = "Merged", 0
                print("检测到：【合并后的账单】", flush=True)
            elif "微信支付账单明细" in cell_A1:
                file_type, skip_rows = "WeChat", 16
                print("检测到：【原始微信文件】", flush=True)
        except Exception as e:
            print(f"分析XLSX文件失败: {e}", flush=True)
            if excel_workbook_object: excel_workbook_object.close()
            return None

    if file_type is None:
        print("错误：无法识别的文件类型。", flush=True)
        print("请确保选择 原始微信账单 / 原始支付宝账单 / 合并后的账dan.xlsx", flush=True)
        if excel_workbook_object: excel_workbook_object.close()
        return None

    # --- 8. 根据文件类型加载数据 ---
    print(f"正在加载数据 (跳过 {skip_rows} 行)...", flush=True)
    try:
        io_source = excel_workbook_object if excel_workbook_object else file_path
        
        if file_type == "WeChat":
            df_source = pd.read_excel(io_source, skiprows=skip_rows, engine='openpyxl')
            df_source.columns = df_source.columns.str.strip()
            df_source['金额(元)'] = pd.to_numeric(df_source['金额(元)'], errors='coerce')
        
        elif file_type == "Alipay":
            df_source = pd.read_csv(io_source, skiprows=skip_rows, encoding=csv_encoding)
            df_source.columns = df_source.columns.str.strip()
            
            amount_col_name = next((col for col in df_source.columns if "金额" in col), None)
            if not amount_col_name:
                raise ValueError("在支付宝文件中找不到任何 '金额' 列。")
                
            df_source.rename(columns={amount_col_name: '金额(元)'}, inplace=True)
            df_source.rename(columns=ALIPAY_MAPPING, inplace=True)
            df_source['金额(元)'] = pd.to_numeric(df_source['金额(元)'], errors='coerce')

        elif file_type == "Merged":
            df_source = pd.read_excel(io_source, skiprows=skip_rows, engine='openpyxl')
            df_source.columns = df_source.columns.str.strip()
            df_source['金额(元)'] = pd.to_numeric(
                df_source['金额(元)'].astype(str).str.replace('¥', '').str.strip(),
                errors='coerce'
            )
            
        if excel_workbook_object:
            excel_workbook_object.close()
            print("... (已关闭 Excel 文件句柄)", flush=True)

        print(f"成功读取 {len(df_source)} 行数据。", flush=True)
        return df_source
        
    except Exception as e:
        print(f"加载数据时出错: {e}", flush=True)
        if excel_workbook_object: excel_workbook_object.close()
        return None

# ******************************************************
# *** 新增：统一的描述处理函数 ***
# ******************************************************
def process_descriptions(df: pd.DataFrame) -> pd.DataFrame:
    """
    按顺序应用所有“描述”列的逻辑：
    1. 设置默认值 (来自 '商品' 或 规则表中的 '描述')
    2. 应用充电桩特殊逻辑
    3. 应用最终的清理规则 (正则表达式)
    
    注意: 此函数假定 '描述_rule', '商品' 和 '备注' (可选) 列已存在于 df 中。
    """
    
    # --- 1. 设置默认值 ---
    # (来自 'c. 重命名规则列' 步骤)
    # 规则表中的'描述' (即'描述_rule') 优先, 其次是 '商品'
    print("...[描述处理] 1/3: 设置默认值 (来自规则或商品)...", flush=True)
    df['描述'] = df['描述'].fillna(df['商品'])

    # --- 2. 应用充电桩特殊逻辑 ---
    # (来自 'd. (硬编码) 处理充电桩' 步骤)
    print("...[描述处理] 2/3: 应用充电桩描述逻辑...", flush=True)
    is_charging = df['商品'].str.contains("自助服务-充电桩", na=False)
    
    charging_desc_split = df.loc[is_charging, '商品'].str.split('/', n=1, expand=True)
    if not charging_desc_split.empty:
        valid_desc = charging_desc_split[1].str.strip()
        # .where() 语句: 如果 valid_desc 有效, 则使用它; 否则, 回退到使用原始的 '商品' 描述
        df.loc[is_charging, '描述'] = valid_desc.where(
            valid_desc.notna() & (valid_desc != ''), 
            df.loc[is_charging, '商品'] # Fallback
        )
    
    # --- 3. 应用最终的清理规则 ---
    # (来自 '*** (最终修复)' 步骤)
    print("...[描述处理] 3/3: 应用最终清理规则 (支付, 订单号, / , ...)...", flush=True)
    
    # 条件1: '描述' 列本身包含 (支付, /, 错误, 订单编号, 或10+位长数字)
    pattern_desc_clear = r"支付|二维码收款|/|错误|订单编号|\d{10,}"
    condition_desc = df['描述'].astype(str).str.contains(
        pattern_desc_clear, 
        na=False, 
        regex=True
    )
    
    # 条件2: '备注' 列包含 (订单编号, 或10+位长数字)
    pattern_memo_clear = r'\d{10,}|订单编号' 
    condition_memo = pd.Series(False, index=df.index) # 默认所有行都 False
    
    if '备注' in df.columns:
        condition_memo = df['备注'].astype(str).str.contains(
            pattern_memo_clear,
            na=False, # na=False (NaN不匹配)
            regex=True
        )
        
    # 合并条件: 只要 描述 匹配 OR 备注 匹配, 就清空 描述
    rows_to_clear = condition_desc | condition_memo
    df.loc[rows_to_clear, '描述'] = ""
    
    print("...描述处理完成。", flush=True)
    return df

# ******************************************************
# *** 核心处理函数 (已重构) ***
# ******************************************************
def process_transactions(df_source, df_rules, start_datetime_obj):
    """
    核心处理函数：清理数据、筛选、并使用向量化应用规则。
    返回: df_final 或 None
    """
    
    # --- 9. 转换日期 ---
    df_source['交易时间'] = df_source['交易时间'].apply(robust_date_converter)
    print("已使用智能转换函数处理 '交易时间' 列。", flush=True)

    # --- 10. 应用时间筛选 ---
    if start_datetime_obj:
        df_source = df_source.loc[
            df_source['交易时间'].notna() & (df_source['交易时间'] >= start_datetime_obj)
        ].copy()
        print(f"应用时间筛选：剩余 {len(df_source)} 行。", flush=True)
    else:
        df_source = df_source.loc[df_source['交易时间'].notna()].copy()
        
    if df_source.empty:
        print("警告：没有符合时间条件的有效数据。", flush=True)
        return pd.DataFrame(columns=TARGET_HEADERS) # 返回空DF

    # --- 11. 筛选 "支出" 且 排除 "已全额退款" 和 "交易关闭" ---
    expenses_condition = (df_source["收/支"] == "支出")
    
    # 默认情况下，所有行都通过状态检查
    status_condition = pd.Series(True, index=df_source.index)
    
    if "当前状态" in df_source.columns:
        # 要排除的状态列表
        excluded_statuses = ["已全额退款", "交易关闭"]
        status_condition = (~df_source["当前状态"].isin(excluded_statuses))
        
    df_expenses = df_source[expenses_condition & status_condition].copy()
    
    print(f"已筛选出 {len(df_expenses)} 条 '支出' (并排除了 '已全额退款' 和 '交易关闭') 记录进行处理...", flush=True)

    if df_expenses.empty:
        print("警告：没有符合条件的支出记录。", flush=True)
        return pd.DataFrame(columns=TARGET_HEADERS) # 返回空DF
        
    # --- 12. (性能优化) 账户清理 (向量化) ---
    print("正在进行性能优化 (向量化处理)...", flush=True)
    try:
        account_map_regex = {
            r'^平安银行信用卡\(4946\).*': '平安银行4946',
            r'^工商银行储蓄卡\(9579\).*': '工商银行'
        }
        df_expenses['支付方式'] = df_expenses['支付方式'].astype(str).replace(account_map_regex, regex=True)
        print("...账户清理完成。", flush=True)
    except Exception as e:
        print(f"!!! 账户清理步骤发生错误: {e} !!!", flush=True)

    # --- 13 & 14. (核心优化) 向量化处理，替代 iterrows() ---
    print("正在进行向量化规则匹配 (Merge)...", flush=True)
    
    df_target = df_expenses.copy()
    
    # a. 准备合并
    df_target['商家(old)'] = df_target['交易对方'].str.strip()
    
    # b. 核心：Left Merge
    df_target = pd.merge(
        df_target, 
        df_rules, 
        on='商家(old)', 
        how='left',
        suffixes=('', '_rule')
    )
    
    # c. 重命名规则列 (注意: '描述' 的默认值设置已移至 process_descriptions)
    df_target.rename(columns={
        '名称_rule': '名称',
        '商家_rule': '商家',
        '描述_rule': '描述', # '描述' 列现在可能来自规则(有值)或(NaN)
        '主类别_rule': '主类别',
        '子类别_rule': '子类别',
        '项目_rule': '项目'
    }, inplace=True)
    
    # '商家' 逻辑: (仍然在这里处理)
    df_target['商家'] = df_target['商家'].fillna(df_target['交易对方'])

    # d. (硬编码) 充电桩非描述逻辑 (分类等)
    print("...正在应用充电桩特殊逻辑 (分类)...", flush=True)
    is_charging_flag = df_target['商品'].str.contains("自助服务-充电桩", na=False)
    df_target.loc[is_charging_flag, ['名称', '主类别', '子类别']] = ['充电', '交通', '加油充电']
    
    # ******************************************************
    # *** (2025-11-03 重构) 统一调用描述处理函数 ***
    df_target = process_descriptions(df_target)
    # ******************************************************

    # e. 格式化标准 Moze 列 (向量化)
    print("...正在格式化目标列...", flush=True)
    df_target['金额'] = pd.to_numeric(df_target['金额(元)'], errors='coerce') * -1
    df_target['日期'] = df_target['交易时间'].dt.strftime('%Y/%m/%d')
    df_target['时间'] = df_target['交易时间'].dt.strftime('%H:%M:%S')
    df_target['_Sort_Date'] = df_target['交易时间'] # 用于排序
    
    df_target['账户'] = df_target['支付方式'] # 已在步骤 12 清理
    df_target['币种'] = "CNY"
    df_target['记录类型'] = "支出"
    df_target['手续费'] = 0
    df_target['折扣'] = 0
    
    # f. 清理 NaNs (来自 merge 失败的规则列)
    fill_na_cols = ['主类别', '子类别', '名称', '项目', '标签', '对象']
    for col in fill_na_cols:
        if col not in df_target.columns:
            df_target[col] = "" # 如果列不存在, 创建空列
    
    df_target[fill_na_cols] = df_target[fill_na_cols].fillna("")
    
    print("...向量化处理完成。", flush=True)
    
    # 只保留最终需要的列
    final_cols = TARGET_HEADERS + ['_Sort_Date']
    df_final = df_target[[col for col in final_cols if col in df_target.columns]]
    
    return df_final


# --- * 3. 主函数 * ---

def main():
    """
    主程序执行流程
    """
    try:
        # 1. 获取输入
        source_file, start_date = get_user_inputs()
        if not source_file:
            # (已在 get_user_inputs 中打印"操作已取消")
            return

        # 2. 检查/创建目标文件夹
        if not TARGET_DIR.exists():
            try:
                TARGET_DIR.mkdir(parents=True, exist_ok=True)
                print(f"已自动创建目标文件夹: {TARGET_DIR}", flush=True)
            except Exception as e:
                print(f"创建文件夹失败: {e}", flush=True)
                return

        # 3. 加载规则
        df_rules = load_rules(RULE_BOOK_PATH)
        if df_rules is None:
            return
            
        # 4. 嗅探和加载数据
        df_source = sniff_and_load_data(source_file)
        if df_source is None:
            print("加载数据失败，程序退出。", flush=True)
            return

        # 5. 核心处理
        df_final = process_transactions(df_source, df_rules, start_date)
        
        if df_final.empty:
            print("没有可导出的数据，程序已完成。", flush=True)
            return

        # 6. 按日期降序排列
        print("正在按日期降序排列 (最新在上)...", flush=True)
        df_final.sort_values(by="_Sort_Date", ascending=False, inplace=True)
        
        # 7. 保存
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        target_path = TARGET_DIR / f"MOZE导入_{timestamp}.csv"
        
        df_final.to_csv(
            target_path, 
            index=False, 
            columns=TARGET_HEADERS, # 确保只保存目标列, _Sort_Date 被丢弃
            encoding='utf-8-sig'
        )
        
        print("-" * 30, flush=True)
        print(f"处理完成！已成功生成文件: '{target_path}'", flush=True)  
        print(f"总共写入了 {len(df_final)} 条记录。", flush=True)

    except Exception as e:
        print(f"\n处理过程中发生未捕获的严重错误: {e}", flush=True)
        print("请检查：", flush=True)
        print("1. 您选择的 Excel/CSV 文件是否是支持的格式？", flush=True)
        print(f"2. 您的 {RULE_BOOK_PATH.name} 格式是否正确？", flush=True)
        print(f"3. 确认您有权限写入 '{TARGET_DIR}' 文件夹。", flush=True)
        # 打印详细的错误堆栈
        print("\n--- 详细错误信息 ---")
        traceback.print_exc()
        print("---------------------\n")


if __name__ == "__main__":
    main()
    # 跨平台暂停
    input("\n...操作已完成，按 [Enter] 键退出。")