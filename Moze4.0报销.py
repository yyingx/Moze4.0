# -*- coding: utf-8 -*-
"""
Moze 报销生成器 v2.0 (极简版)
特性: 
1. 输入 yymmdd (如 250101)
2. 交互式填写出差地点 -> 自动写入[描述]
3. 自动归档至 Moze4.0_Import
"""

import pandas as pd
import datetime
import tkinter as tk
from tkinter import simpledialog, messagebox
from pathlib import Path

# --- 核心配置 ---
CFG = {
    'MEAL': 40,      # 餐补
    'TRANS': 20,     # 交补
    'HOTEL': 140,    # 住宿标准
    'ACC': '钱包',   # 账户
    'MERCHANT': '天之逸',
    'TAG': '#差旅费报销'
}

COLUMNS = ['账户', '币种', '记录类型', '主类别', '子类别', '金额', '手续费', '折扣', '名称', '商家', '日期', '时间', '项目', '描述', '标签', '对象']

class MozeReimburse:
    def __init__(self):
        self.rows = []
        self.root = tk.Tk()
        self.root.withdraw() # 隐藏主窗口

    def parse_dates(self, date_str):
        """解析 yymmdd 或 yymmdd-yymmdd"""
        try:
            parts = date_str.split('-')
            fmt = "%y%m%d" # 关键修改：两位年份
            start = datetime.datetime.strptime(parts[0].strip(), fmt)
            end = datetime.datetime.strptime(parts[-1].strip(), fmt) # 若无杠，start=end
            return [start + datetime.timedelta(days=i) for i in range((end - start).days + 1)]
        except:
            return None

    def add(self, date_obj, name, amt, loc="", offset=0):
        """生成双向分录: 收入(+), 报账(-)"""
        if amt <= 0: return
        d_str = date_obj.strftime('%Y-%m-%d') # 输出仍保持标准格式兼容App
        base = datetime.datetime(2025, 1, 1, 17, 30, 0) + datetime.timedelta(minutes=offset)
        
        # 1. 收入
        self.rows.append({
            '账户': CFG['ACC'], '币种': 'CNY', '记录类型': '收入', 
            '主类别': '收入', '子类别': '福利补贴', '金额': amt,
            '手续费': 0, '折扣': 0, '名称': name, '商家': CFG['MERCHANT'],
            '日期': d_str, '时间': base.strftime('%H:%M:%S'), '项目': '工作',
            '描述': loc, '标签': '', '对象': '' # 描述填入地点
        })
        # 2. 应收冲抵
        self.rows.append({
            '账户': CFG['ACC'], '币种': 'CNY', '记录类型': '应收款项', 
            '主类别': '应收款项', '子类别': '报账', '金额': -amt,
            '手续费': 0, '折扣': 0, '名称': name, '商家': '',
            '日期': d_str, '时间': (base + datetime.timedelta(seconds=1)).strftime('%H:%M:%S'), '项目': '',
            '描述': loc, '标签': CFG['TAG'], '对象': CFG['MERCHANT'] # 描述填入地点
        })

    def run(self):
        while True:
            mode = simpledialog.askinteger("Moze生成器", "1. 餐补/交补 (批量)\n2. 住宿补贴 (差额)\n3. 导出", parent=self.root)
            if not mode: break

            if mode == 1: # 批量模式
                d_in = simpledialog.askstring("步骤1/3", "日期 (如 250901-250905):")
                if not d_in: continue
                dates = self.parse_dates(d_in)
                if not dates: 
                    messagebox.showerror("错", "格式应为 yymmdd")
                    continue
                
                loc = simpledialog.askstring("步骤2/3", "出差地点 (填入描述):", initialvalue="武汉")
                
                # 询问补贴
                do_meal = messagebox.askyesno("步骤3/3", f"加【餐补】{CFG['MEAL']}元?")
                do_trans = messagebox.askyesno("步骤3/3", f"加【交补】{CFG['TRANS']}元?")

                for d in dates:
                    if do_meal: self.add(d, "餐费补贴", CFG['MEAL'], loc, 0)
                    if do_trans: self.add(d, "交通补贴", CFG['TRANS'], loc, 10)
                messagebox.showinfo("OK", f"已添加 {len(dates)} 天记录")

            elif mode == 2: # 住宿模式
                d_in = simpledialog.askstring("住宿", "入住日期 (yymmdd):")
                if not d_in: continue
                dates = self.parse_dates(d_in)
                if not dates: continue
                
                loc = simpledialog.askstring("住宿", "出差地点:", initialvalue="武汉")
                cost = simpledialog.askfloat("住宿", f"标准 {CFG['HOTEL']} 元\n输入实际花费:")
                
                if cost:
                    sub = CFG['HOTEL'] - cost
                    if sub > 0:
                        self.add(dates[0], "住宿补贴", sub, loc, 120)
                        messagebox.showinfo("OK", f"补贴 {sub} 元")
                    else:
                        messagebox.showwarning("无补贴", "超标或刚好")

            elif mode == 3: # 导出
                if not self.rows: break
                # 保存逻辑
                path = Path(__file__).parent / "Moze4.0_Import"
                path.mkdir(exist_ok=True)
                f_name = path / f"Moze_报销_{datetime.datetime.now().strftime('%y%m%d_%H%M%S')}.csv"
                pd.DataFrame(self.rows, columns=COLUMNS).to_csv(f_name, index=False, encoding='utf-8-sig')
                messagebox.showinfo("完成", f"已保存:\n{f_name.name}")
                break
        self.root.destroy()

if __name__ == "__main__":
    MozeReimburse().run()