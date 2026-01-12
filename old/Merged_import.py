import pandas as pd
import os
import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog

# --- * 智能日期转换函数 * ---
def robust_date_converter(x):
    """
    智能处理混合类型的日期列 (字符串, Excel浮点数, 或 datetime)
    """
    if isinstance(x, (datetime.datetime, datetime.date)):
        # 如果已经是日期时间对象, 直接返回
        return pd.to_datetime(x)
    elif isinstance(x, (int, float)):
        # 如果是数字 (可能是Excel的浮点数时间戳)
        try:
            # 1899-12-30 是 Excel 的 'origin'
            return pd.to_datetime(x, unit='D', origin='1899-12-30')
        except Exception:
            return pd.NaT  # 转换失败
    elif isinstance(x, str):
        # 如果是字符串, 正常解析
        return pd.to_datetime(x, errors='coerce')
    else:
        # 其他类型, 视为无效
        return pd.NaT

# --- 1. 文件选择对话框 ---
print("正在打开文件选择对话框，请选择您的 【合并后的账单.xlsx】 文件...")
root = tk.Tk()
root.withdraw()  # 隐藏主窗口

# 弹出文件选择器
source_excel_file = filedialog.askopenfilename(
    title="请选择您要处理的 【合并后的账单.xlsx】",
    filetypes=(("Excel 文件", "*.xlsx"), ("所有文件", "*.*"))
)

# 检查用户是否选择了文件
if not source_excel_file:
    print("操作已取消：您没有选择任何文件。")
    root.destroy()
    exit()
    
print(f"您选择的源文件是: {source_excel_file}")

# --- 2. 简化日期筛选对话框 ---
start_time_str = simpledialog.askstring(
    "筛选条件 (按日期)", 
    "请输入起始日期（包含当天及之后）\n\n"
    "**格式： yymmdd**\n"
    "（例如: 251001  代表 2025年10月1日）\n\n"
    "*** 如果留空，将导入所有时间的数据 ***",
    parent=root
)
root.destroy()

start_datetime_obj = None

# 检查用户是否输入了内容
if start_time_str and start_time_str.strip() != "":
    try:
        # 使用 %y%m%d 格式解析输入
        start_datetime_obj = datetime.datetime.strptime(start_time_str.strip(), "%y%m%d")
        print(f"筛选条件：将只导入 {start_datetime_obj.strftime('%Y-%m-%d')} 00:00:00 之后的数据。")
    except ValueError:
        print(f"错误：日期格式不正确！'{start_time_str}' 不是 yymmdd 格式。")
        print("程序已退出。")
        exit()
else:
    print("提示：未输入起始时间，将导入文件中的所有数据。")


# --- 3. 定义目标文件夹和文件名 ---
target_dir = r"D:\OneDrive\桌面\Moze4.0\Moze4.0_Import"

# 检查文件夹是否存在，如果不存在则自动创建
if not os.path.exists(target_dir):
    try:
        os.makedirs(target_dir)
        print(f"已自动创建目标文件夹: {target_dir}")
    except Exception as e:
        print(f"创建文件夹失败: {e}")
        exit()

# 生成文件名
now = datetime.datetime.now()
timestamp_str = now.strftime("%Y%m%d_%H%M%S")  
file_name = f"MOZE导入_{timestamp_str}.csv" 
target_csv_file = os.path.join(target_dir, file_name) 
# ---

# --- 4. 目标 CSV 抬头 ---
target_headers = [
    "账户", "币种", "记录类型", "主类别", "子类别", "金额", 
    "手续费", "折扣", "名称", "商家", "日期", "时间", 
    "项目", "描述", "标签", "对象"
]

print(f"将要生成的目标文件: '{target_csv_file}'") 

try:
    # 1. 读取 Excel
    df_source = pd.read_excel(
        source_excel_file
    )
    print(f"成功读取 {len(df_source)} 行数据。")
    
    # 2. 使用我们的智能转换函数
    df_source['交易时间'] = df_source['交易时间'].apply(robust_date_converter)
    print("已使用智能转换函数处理 '交易时间' 列。")

    # 3. 应用时间筛选
    if start_datetime_obj:
        original_count = len(df_source)
        df_source = df_source.loc[
            df_source['交易时间'].notna() & (df_source['交易时间'] >= start_datetime_obj)
        ].copy()
        
        filtered_count = len(df_source)
        print(f"应用时间筛选：从 {original_count} 行刷减到 {filtered_count} 行。")
    else:
        df_source = df_source.loc[df_source['交易时间'].notna()].copy()


    # 4. *** 修正：筛选 "支出" 且 排除 "已全额退款" ***
    expenses_condition = (df_source["收/支"] == "支出")
    refund_condition = (df_source["当前状态"] != "已全额退款")
    
    df_expenses = df_source[expenses_condition & refund_condition].copy()
    
    print(f"已筛选出 {len(df_expenses)} 条 '支出' (且未退款) 记录进行处理...")

    # 5. 准备列表
    processed_rows = []

    # 6. 遍历 df_expenses
    for index, row in df_expenses.iterrows():
            
            new_row = {}

            # a. 映射
            new_row["账户"] = row.get("支付方式") 
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
                # *** 新增：添加一个隐藏的排序用列 ***
                new_row["_Sort_Date"] = source_datetime 
            else:
                new_row["日期"] = "日期格式无法解析"
                new_row["时间"] = "时间格式无法解析"
                new_row["_Sort_Date"] = pd.NaT # 确保每行都有

            # d. 映射规则 (自动分类)
            
            # 先获取源数据
            merchant_name = str(row.get("交易对方", ""))
            item_description = str(row.get("商品", ""))
            
            # --- 默认值 ---
            new_row["主类别"] = "" 
            new_row["子类别"] = ""
            new_row["名称"] = ""
            new_row["商家"] = merchant_name
            new_row["描述"] = item_description
            new_row["手续费"] = 0
            new_row["折扣"] = 0
            
            # --- 应用您的新规则 ---
            
            if "自助服务平台" in merchant_name:
                new_row["名称"] = "充电"
                new_row["主类别"] = "交通"
                new_row["子类别"] = "加油充电"
            
            elif merchant_name == "零食顽家":
                new_row["名称"] = ""
                new_row["主类别"] = "饮食"
                new_row["子类别"] = "零食"
                
            elif "蔬菜" in merchant_name:
                new_row["名称"] = "蔬菜"
                new_row["主类别"] = "饮食"
                new_row["子类别"] = "食材"
            
            processed_rows.append(new_row)

    # 7. 转换 (额外加入 _Sort_Date 列)
    df_target = pd.DataFrame(processed_rows, columns=target_headers + ["_Sort_Date"])

    # 8. 替换 (账户字典 - 正则表达式)
    
    # (字典) 精确替换
    account_map_exact = {
        # "零钱": "微信零钱" # 示例
    }
    if account_map_exact: # 只有在字典不为空时才运行
        df_target['账户'] = df_target['账户'].replace(account_map_exact)
        print(f"已根据字典(精确)替换 '账户' 列内容。")

    # (Regex) 模糊/包含替换
    account_map_regex = {
        r'^平安银行信用卡\(4946\).*': '平安银行4946',
        r'^工商银行储蓄卡\(9579\).*': '工商银行'
    }
    df_target['账户'] = df_target['账户'].replace(account_map_regex, regex=True)
    print(f"已根据正则表达式(模糊)替换 '账户' 列内容。")


    # 8.5. 替换 (商家字典 - 正则表达式)
    
    # (字典) 精确替换
    merchant_map_exact = {
        "武汉市洪山区三镇民生甜食馆": "三镇民生甜食馆",
        "王记育林商贸有限公司": "康福路农副产品批发市场" 
    }
    df_target['商家'] = df_target['商家'].replace(merchant_map_exact)
    print(f"已根据字典(精确)替换 '商家' 列内容。")
    
    # (Regex) 模糊/包含替换
    merchant_map_regex = {
        r'.*(兰州拉面|牛肉面).*': '兰州拉面',
        r'^(康晶食品|王记育林商贸有限公司)$': '康福路农副产品批发市场'
    }
    df_target['商家'] = df_target['商家'].replace(merchant_map_regex, regex=True)
    print(f"已根据正则表达式(模糊)替换 '商家' 列内容。")

    # 8.6. 按日期降序排列
    print("正在按日期降序排列 (最新在上)...")
    # (使用我们隐藏的 _Sort_Date 列进行精确排序)
    df_target.sort_values(by="_Sort_Date", ascending=False, inplace=True)


    # 9. 保存 (只保存 target_headers, 排除 _Sort_Date)
    df_target.to_csv(target_csv_file, index=False, columns=target_headers, encoding='utf-8-sig')

    print("-" * 30)
    print(f"处理完成！已成功生成文件: '{target_csv_file}'") 
    print(f"总共写入了 {len(df_target)} 条记录。")

except Exception as e:
    print(f"\n处理过程中发生错误: {e}")
    print("请检查：")
    print("1. 您选择的 Excel 文件是否是正确的 【合并后的账单】 格式？")
    print(f"3. 确认您有权限写入 '{target_dir}' 文件夹。")

# 在最后暂停，以便用户可以看到输出信息
os.system("pause")