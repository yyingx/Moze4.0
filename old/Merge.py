import pandas as pd
import os
import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog

# --- 1. 初始化 Tkinter ---
print("正在启动交互式合并工具...")
root = tk.Tk()
root.withdraw()  # 隐藏主窗口

# --- 2. 选择基础文件 (例如 ALL.xlsx) ---
base_file_path = filedialog.askopenfilename(
    title="第1步: 请选择您的 基础文件 (例如 ALL.xlsx)",
    filetypes=(("Excel/CSV 文件", "*.xlsx;*.csv"), ("所有文件", "*.*"))
)
if not base_file_path:
    print("操作已取消。")
    exit()
print(f"基础文件: {base_file_path}")

# 询问跳过几行
base_skip_rows = simpledialog.askinteger(
    "基础文件",
    f"要跳过几行才能读取 {os.path.basename(base_file_path)} 的抬头?\n\n(例如: 微信/ALL.xlsx 是 16)",
    initialvalue=16
)
if base_skip_rows is None:
    print("操作已取消。")
    exit()

# --- 3. 选择新数据文件 (例如 支付宝.csv) ---
new_data_file_path = filedialog.askopenfilename(
    title="第2步: 请选择要合并的 新数据 (例如 支付宝.csv)",
    filetypes=(("Excel/CSV 文件", "*.xlsx;*.csv"), ("所有文件", "*.*"))
)
if not new_data_file_path:
    print("操作已取消。")
    exit()
print(f"新数据文件: {new_data_file_path}")

# 询问跳过几行
new_skip_rows = simpledialog.askinteger(
    "新数据文件",
    f"要跳过几行才能读取 {os.path.basename(new_data_file_path)} 的抬头?\n\n(例如: 支付宝.csv 是 24)",
    initialvalue=24
)
if new_skip_rows is None:
    print("操作已取消。")
    exit()

root.destroy() # 关闭Tkinter

# --- 4. 动态生成输出文件名 ---
now = datetime.datetime.now()
timestamp_str = now.strftime("%Y%m%d_%H%M%S")
output_file = f"合并后的账单_{timestamp_str}.xlsx"

# --- 5. 定义支付宝的映射规则 (基于我们上次的确认) ---
alipay_mapping = {
    '交易时间': '交易时间',
    '交易分类': '交易类型',
    '交易对方': '交易对方',
    '商品说明': '商品',
    '收/支': '收/支',
    '金额': '金额(元)',      # 关键映射
    '收/付款方式': '支付方式', # 关键映射
    '交易状态': '当前状态',    # 关键映射
    '交易订单号': '交易单号',  # 关键映射
    '商家订单号': '商户单号',  # 关键映射
    '备注': '备注'
}

# --- 6. 定义一个通用的文件读取函数 (自动处理CSV/Excel) ---
def read_bill_file(file_path, skip_rows):
    if file_path.endswith('.csv'):
        try:
            # 尝试 UTF-8
            return pd.read_csv(file_path, skiprows=skip_rows)
        except UnicodeDecodeError:
            # 失败则尝试 GBK (支付宝常用)
            print(f"检测到 GBK 编码: {file_path}")
            return pd.read_csv(file_path, skiprows=skip_rows, encoding='GBK')
    elif file_path.endswith('.xlsx'):
        return pd.read_excel(file_path, skiprows=skip_rows)
    else:
        raise ValueError(f"不支持的文件格式: {file_path}")

# --- 7. 执行合并 ---
try:
    # 1. 读取基础文件
    df_base = read_bill_file(base_file_path, base_skip_rows)
    # 清理列名空白, 并获取目标抬头
    df_base.columns = df_base.columns.str.strip()
    target_headers = df_base.columns.to_list()
    print(f"成功读取 '基础文件'，共 {len(df_base)} 行数据。")
    print(f"目标抬头为: {target_headers}")

    # 2. 读取新数据文件
    df_new = read_bill_file(new_data_file_path, new_skip_rows)
    df_new.columns = df_new.columns.str.strip() # 清理列名空白
    print(f"成功读取 '新数据文件'，共 {len(df_new)} 行新数据。")

    # 3. 假设新文件是支付宝格式, 重命名它的列
    df_new_renamed = df_new.rename(columns=alipay_mapping)

    # 4. 确保新数据只包含目标列, 且顺序一致
    # (这将自动丢弃 '对方账号', 'Unnamed: 12' 等多余的列)
    df_new_transformed = df_new_renamed[target_headers]

    # 5. 合并两个数据表
    df_merged = pd.concat([df_base, df_new_transformed], ignore_index=True)
    print(f"合并完成。合并后总行数: {len(df_merged)}")

    # 6. 智能去重 (基于 '交易单号'，保留第一个出现的)
    original_rows = len(df_merged)
    # (我们还需要清理单号中的制表符)
    df_merged['交易单号'] = df_merged['交易单号'].astype(str).str.strip()
    
    df_merged.drop_duplicates(subset=['交易单号'], keep='first', inplace=True)
    final_rows = len(df_merged)
    
    if original_rows > final_rows:
        print(f"去重完成：已删除 {original_rows - final_rows} 个重复条目。")
    print(f"最终总行数: {final_rows}")

    # 7. 保存到新的 Excel 文件
    df_merged.to_excel(output_file, index=False)
    
    print("\n" + "="*30)
    print("✨ 操作成功! ✨")
    print(f"已将两个文件合并，并保存为新文件:")
    print(output_file)
    print("您可以将此新文件作为下次合并的 '基础文件'。")

except FileNotFoundError as e:
    print(f"错误：找不到文件 {e.filename}。")
except KeyError as e:
    print(f"\n--- 错误：字段映射失败 ---")
    print(f"无法在 '新数据文件' 中找到预期的列: {e}")
    print("这可能意味着：")
    print("  1. 您选择的新文件不是 支付宝 格式。")
    print("  2. 支付宝 的账单格式更新了。")
    print("  3. 您的 '跳过行数' 输入错误。")
except Exception as e:
    print(f"合并过程中发生错误: {e}")

# 在最后暂停，以便用户可以看到输出信息
os.system("pause")
