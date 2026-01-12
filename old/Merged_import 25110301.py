import pandas as pd
import os
import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog

# --- * 智能日期转换函数 * ---
def robust_date_converter(x):
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

# --- 1. 文件选择对话框 (数据文件) ---
print("正在打开文件选择对话框...")
root = tk.Tk()
root.withdraw()  # 隐藏主窗口

print("第1步: 请选择您的 【合并后的账单.xlsx】 文件...")
source_excel_file = filedialog.askopenfilename(
    title="第1步: 请选择您要处理的 【合并后的账单.xlsx】",
    filetypes=(("Excel 文件", "*.xlsx"), ("所有文件", "*.*"))
)
if not source_excel_file:
    print("操作已取消。")
    root.destroy()
    exit()
print(f"您选择的源文件是: {source_excel_file}")

# --- 2. 硬编码规则文件路径 ---
rule_book_path = r"D:\OneDrive\桌面\Moze4.0\Moze Dict.csv"
print(f"将自动加载规则文件: {rule_book_path}")

# 检查规则文件是否存在
if not os.path.exists(rule_book_path):
    print(f"错误：找不到规则文件！")
    print(f"请确保您的规则文件位于: {rule_book_path}")
    root.destroy()
    os.system("pause")
    exit()

# --- 3. 简化日期筛选对话框 ---
start_time_str = simpledialog.askstring(
    "第2步: 筛选条件 (按日期)",  # (已改为第2步)
    "请输入起始日期（包含当天及之后）\n\n"
    "**格式： yymmdd**\n"
    "（例如: 251001  代表 2025年10月1日）\n\n"
    "*** 如果留空，将导入所有时间的数据 ***",
    parent=root
)
root.destroy() # 关闭所有对话框

start_datetime_obj = None
if start_time_str and start_time_str.strip() != "":
    try:
        start_datetime_obj = datetime.datetime.strptime(start_time_str.strip(), "%y%m%d")
        print(f"筛选条件：将只导入 {start_datetime_obj.strftime('%Y-%m-%d')} 00:00:00 之后的数据。")
    except ValueError:
        print(f"错误：日期格式不正确！'{start_time_str}' 不是 yymmdd 格式。")
        print("程序已退出。")
        os.system("pause")
        exit()
else:
    print("提示：未输入起始时间，将导入文件中的所有数据。")


# --- 4. 定义目标文件夹和文件名 ---
target_dir = r"D:\OneDrive\桌面\Moze4.0\Moze4.0_Import"

if not os.path.exists(target_dir):
    try:
        os.makedirs(target_dir)
        print(f"已自动创建目标文件夹: {target_dir}")
    except Exception as e:
        print(f"创建文件夹失败: {e}")
        os.system("pause")
        exit()

now = datetime.datetime.now()
timestamp_str = now.strftime("%Y%m%d_%H%M%S")  
file_name = f"MOZE导入_{timestamp_str}.csv" 
target_csv_file = os.path.join(target_dir, file_name) 

# --- 5. 目标 CSV 抬头 ---
target_headers = [
    "账户", "币种", "记录类型", "主类别", "子类别", "金额", 
    "手续费", "折扣", "名称", "商家", "日期", "时间", 
    "项目", "描述", "标签", "对象"
]

print(f"将要生成的目标文件: '{target_csv_file}'") 

try:
    # --- 6. 读取规则.csv 文件 ---
    print(f"正在读取规则文件: {rule_book_path}")
    
    # 使用 pd.read_csv() 来快速加载
    df_rules = pd.read_csv(rule_book_path, encoding='utf-8-sig')
    
    # 确保 "商家(old)" 列存在
    if '商家(old)' not in df_rules.columns:
        print(f"错误：规则文件 {rule_book_path} 中缺少 '商家(old)' 列！")
        os.system("pause")
        exit()
        
    # 清理 商家(old) 名称中的空白
    df_rules['商家(old)'] = df_rules['商家(old)'].str.strip()
    
    # 自动删除 '商家(old)' 列中的重复项, 保留最后一个
    df_rules.drop_duplicates(subset=['商家(old)'], keep='last', inplace=True)
    
    # (我们使用 .fillna('') 确保Excel中的空白单元格被读为空字符串)
    rule_book_dict = df_rules.set_index('商家(old)').fillna('').to_dict('index')
    print(f"成功加载 {len(rule_book_dict)} 条[唯一]分类规则。")
    
    # --- 7. (已删除) 内置的【商家清理】规则 (不再需要) ---
    print("商家清理规则已由 'Moze Dict.csv' 的 '商家(old)' 列接管。")

    # --- 8. 读取数据文件 ---
    df_source = pd.read_excel(source_excel_file)
    print(f"成功读取 {len(df_source)} 行数据。")
    
    # 9. 转换日期
    df_source['交易时间'] = df_source['交易时间'].apply(robust_date_converter)
    print("已使用智能转换函数处理 '交易时间' 列。")

    # 10. 应用时间筛选
    if start_datetime_obj:
        original_count = len(df_source)
        df_source = df_source.loc[
            df_source['交易时间'].notna() & (df_source['交易时间'] >= start_datetime_obj)
        ].copy()
        
        filtered_count = len(df_source)
        print(f"应用时间筛选：从 {original_count} 行刷减到 {filtered_count} 行。")
    else:
        df_source = df_source.loc[df_source['交易时间'].notna()].copy()


    # 11. 筛选 "支出" 且 排除 "已全额退款"
    expenses_condition = (df_source["收/支"] == "支出")
    refund_condition = (df_source["当前状态"] != "已全额退款")
    
    df_expenses = df_source[expenses_condition & refund_condition].copy()
    
    print(f"已筛选出 {len(df_expenses)} 条 '支出' (且未退款) 记录进行处理...")
    
    # -----------------------------------------------------------------
    # --- 12. (性能优化) 账户清理 (向量化) *** ---
    # -----------------------------------------------------------------
    print("正在进行性能优化 (向量化处理)...")
    
    account_map_regex = {
        r'^平安银行信用卡\(4946\).*': '平安银行4946',
        r'^工商银行储蓄卡\(9579\).*': '工商银行'
    }
    # (我们先对 *所有* 支出行一次性应用账户替换)
    df_expenses['支付方式'] = df_expenses['支付方式'].astype(str).replace(account_map_regex, regex=True)
    
    # -----------------------------------------------------------------

    # 13. 准备列表
    processed_rows = []

    # 14. 遍历 df_expenses
    print("正在循环处理每一行数据...")
    for index, row in df_expenses.iterrows():
            
            new_row = {}

            # a. 映射
            new_row["账户"] = row.get("支付方式") # <-- 读取已清理的账户
            new_row["币种"] = "CNY"
            new_row["记录类型"] = "支出" 
            
            # b. 金额处理
            amount_str = str(row.get("金额(元)", "")).replace("¥", "").strip()
            numeric_amount = pd.to_numeric(amount_str, errors='coerce')
            
            new_row["金额"] = numeric_amount * -1 # 支出为负
            
            # c. 手动构建日期和时间字符串
            source_datetime = row.get("交易时间") 
            if pd.notna(source_datetime):
                new_row["日期"] = f"{source_datetime.year}/{source_datetime.month}/{source_datetime.day}"
                new_row["时间"] = f"{source_datetime.hour:02}:{source_datetime.minute:02}:{source_datetime.second:02}"
                new_row["_Sort_Date"] = source_datetime 
            else:
                new_row["日期"] = "日期格式无法解析"
                new_row["时间"] = "时间格式无法解析"
                new_row["_Sort_Date"] = pd.NaT 

            # d. *** 修正：已修复 "项目" Bug ***
            
            # 先获取源数据
            merchant_name = str(row.get("交易对方", "")).strip()
            item_description = str(row.get("商品", "")).strip()
            
            # --- 默认值 ---
            new_row["主类别"] = "" 
            new_row["子类别"] = ""
            new_row["名称"] = ""
            new_row["商家"] = merchant_name        # 默认 = 交易对方 (raw)
            new_row["描述"] = item_description  # 默认 = 商品
            new_row["项目"] = ""                # 默认 = 空
            new_row["手续费"] = 0
            new_row["折扣"] = 0
            
            # --- 步骤 1: (硬编码的特殊逻辑 - 充电) ---
            # (检查 '商品' 列)
            is_charging_product = "自助服务-充电桩" in item_description
            
            if is_charging_product:
                
                # *** 修正：先从规则库里查找 "项目" ***
                if merchant_name in rule_book_dict:
                    rule = rule_book_dict[merchant_name]
                    new_row["项目"] = rule.get('项目', '') # <-- 修复 Bug
                
                # (然后应用特殊规则)
                new_row["名称"] = "充电"
                new_row["主类别"] = "交通"
                new_row["子类别"] = "加油充电"
                
                # (检查 '商品' 列是否有斜杠)
                if "/" in item_description:
                    parts = item_description.split('/', 1)
                    if len(parts) == 2 and parts[1].strip() != "":
                        new_row["描述"] = parts[1].strip() # 描述 = G371431-2 (变量)

            # --- 步骤 2: (新的规则库) ---
            # (仅在 '充电桩' *不* 匹配时才运行)
            # (用 *原始* 商家名称 `merchant_name` 去匹配 `商家(old)` 键)
            elif merchant_name in rule_book_dict:
                rule = rule_book_dict[merchant_name]
                
                # 应用所有规则
                new_row["项目"] = rule.get('项目', '')
                new_row["主类别"] = rule.get('主类别', '')
                new_row["子类别"] = rule.get('子类别', '')
                
                rule_name = rule.get('名称')
                rule_desc = rule.get('描述')
                rule_clean_merchant = rule.get('商家') # <-- 新的干净商家名
                
                if rule_name: 
                    new_row["名称"] = rule_name
                
                if rule_desc: 
                    new_row["描述"] = rule_desc
                # (否则, '描述' 保持默认的 item_description - "包粑" Bug 已修复)
                
                if rule_clean_merchant:
                    new_row["商家"] = rule_clean_merchant # <-- 覆盖为新的干净商家名
                # (否则, '商家' 保持默认的 merchant_name (raw))
            
            processed_rows.append(new_row)

    # 15. 转换
    df_target = pd.DataFrame(processed_rows, columns=target_headers + ["_Sort_Date"])

    # 16. 替换 (账户字典 - *注意: 这一步已被移到循环 *前* *)
    print(f"账户名称清理已在循环前完成。")

    # 17. 替换 (商家字典 - *注意: 这一步已被移入规则文件*)
    print(f"商家名称清理已由 'Moze Dict.csv' 接管。")

    # 18. 按日期降序排列
    print("正在按日期降序排列 (最新在上)...")
    df_target.sort_values(by="_Sort_Date", ascending=False, inplace=True)


    # 19. 保存 (只保存 target_headers, 排除 _Sort_Date)
    df_target.to_csv(target_csv_file, index=False, columns=target_headers, encoding='utf-8-sig')

    print("-" * 30)
    print(f"处理完成！已成功生成文件: '{target_csv_file}'") 
    print(f"总共写入了 {len(df_target)} 条记录。")

except Exception as e:
    print(f"\n处理过程中发生错误: {e}")
    print("请检查：")
    print("1. 您选择的 Excel 文件是否是正确的 【合并后的账单】 格式？")
    # *** 修正：已将 'print' 语句合并到一行 ***
    print("2. 您选择的 【规则文件】 格式是否正确 (必须包含 '商家(old)' 列, 且 '商家(old)' 列内容唯一)？")
    print(f"3. 确认您有权限写入 '{target_dir}' 文件夹。")

# 在最后暂停，以便用户可以看到输出信息
os.system("pause")