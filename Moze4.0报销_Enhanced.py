# -*- coding: utf-8 -*-
"""
Moze 报销生成器 v2.1 (增强版)
特性: 
1. 输入 yymmdd (如 250101)
2. 交互式填写出差地点 -> 自动写入[描述]
3. 自动归档至 Moze4.0_Import
4. 【新增】预览功能 - 导出前查看汇总
5. 【新增】撤销功能 - 删除最后一批记录
6. 【新增】自定义金额 - 支持手动输入特殊补贴
7. 【优化】更友好的界面提示
"""

import pandas as pd
import datetime
import tkinter as tk
from tkinter import simpledialog, messagebox, scrolledtext
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
        self.batch_sizes = []  # 记录每批添加的记录数，用于撤销
        self.root = tk.Tk()
        self.root.withdraw() # 隐藏主窗口

    def parse_dates(self, date_str):
        """解析 yymmdd 或 yymmdd-yymmdd"""
        try:
            parts = date_str.split('-')
            fmt = "%y%m%d" # 两位年份
            start = datetime.datetime.strptime(parts[0].strip(), fmt)
            end = datetime.datetime.strptime(parts[-1].strip(), fmt) # 若无杠，start=end
            return [start + datetime.timedelta(days=i) for i in range((end - start).days + 1)]
        except:
            return None

    def add(self, date_obj, name, amt, loc="", offset=0):
        """生成双向分录: 收入(+), 报账(-)"""
        if amt <= 0: return 0
        d_str = date_obj.strftime('%Y/%m/%d') # 修改为 YYYY/MM/DD 格式
        base = datetime.datetime(2025, 1, 1, 17, 30, 0) + datetime.timedelta(minutes=offset)
        
        # 1. 收入
        self.rows.append({
            '账户': CFG['ACC'], '币种': 'CNY', '记录类型': '收入', 
            '主类别': '收入', '子类别': '福利补贴', '金额': amt,
            '手续费': 0, '折扣': 0, '名称': name, '商家': CFG['MERCHANT'],
            '日期': d_str, '时间': base.strftime('%H:%M:%S'), '项目': '工作',
            '描述': loc, '标签': '', '对象': ''
        })
        # 2. 应收冲抵
        self.rows.append({
            '账户': CFG['ACC'], '币种': 'CNY', '记录类型': '应收款项', 
            '主类别': '应收款项', '子类别': '报账', '金额': -amt,
            '手续费': 0, '折扣': 0, '名称': name, '商家': '',
            '日期': d_str, '时间': (base + datetime.timedelta(seconds=1)).strftime('%H:%M:%S'), '项目': '',
            '描述': loc, '标签': CFG['TAG'], '对象': CFG['MERCHANT']
        })
        return 2  # 返回添加的记录数

    def show_preview(self):
        """显示当前记录预览"""
        if not self.rows:
            messagebox.showinfo("预览", "暂无记录")
            return
        
        # 创建预览窗口
        preview_win = tk.Toplevel(self.root)
        preview_win.title("报销记录预览")
        preview_win.geometry("800x500")
        
        # 创建滚动文本框
        text_area = scrolledtext.ScrolledText(preview_win, wrap=tk.WORD, width=100, height=25, font=("Courier", 10))
        text_area.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        
        # 统计信息
        df = pd.DataFrame(self.rows)
        total_income = df[df['记录类型'] == '收入']['金额'].sum()
        total_records = len(self.rows)
        
        # 按日期分组统计
        summary = df.groupby(['日期', '名称']).agg({
            '金额': lambda x: x[x > 0].sum()  # 只统计收入（正数）
        }).reset_index()
        
        # 生成预览文本
        preview_text = f"{'='*60}\n"
        preview_text += f"  📊 报销记录预览\n"
        preview_text += f"{'='*60}\n\n"
        preview_text += f"📌 总记录数: {total_records} 条 (收入+应收各占一半)\n"
        preview_text += f"💰 总补贴金额: ¥{total_income:.2f}\n\n"
        preview_text += f"{'='*60}\n"
        preview_text += f"  📅 明细清单\n"
        preview_text += f"{'='*60}\n\n"
        
        # 按日期分组显示
        current_date = None
        for _, row in df[df['记录类型'] == '收入'].iterrows():
            date = row['日期']
            if date != current_date:
                if current_date is not None:
                    preview_text += "\n"
                preview_text += f"📅 {date}\n"
                preview_text += f"{'-'*60}\n"
                current_date = date
            
            preview_text += f"  • {row['名称']:<12} ¥{row['金额']:>8.2f}  [{row['描述']}]\n"
        
        preview_text += f"\n{'='*60}\n"
        preview_text += f"  💡 提示\n"
        preview_text += f"{'='*60}\n"
        preview_text += f"• 每条补贴对应 2 条记录（收入+应收冲抵）\n"
        preview_text += f"• 所有记录带标签: {CFG['TAG']}\n"
        preview_text += f"• 导出路径: Moze4.0_Import/\n"
        
        text_area.insert(tk.END, preview_text)
        text_area.config(state=tk.DISABLED)  # 设置为只读
        
        # 添加关闭按钮
        close_btn = tk.Button(preview_win, text="关闭", command=preview_win.destroy, font=("Arial", 12))
        close_btn.pack(pady=10)

    def undo_last_batch(self):
        """撤销最后一批记录"""
        if not self.batch_sizes:
            messagebox.showinfo("撤销", "没有可撤销的记录")
            return
        
        last_batch_size = self.batch_sizes.pop()
        removed = self.rows[-last_batch_size:]
        self.rows = self.rows[:-last_batch_size]
        
        # 统计撤销信息
        income_records = [r for r in removed if r['记录类型'] == '收入']
        total = sum(r['金额'] for r in income_records)
        
        messagebox.showinfo("撤销成功", f"已删除 {len(income_records)} 条补贴记录\n总金额: ¥{total:.2f}")

    def run(self):
        while True:
            menu_text = "=" * 40 + "\n"
            menu_text += "  🎯 Moze 报销生成器 v2.1\n"
            menu_text += "=" * 40 + "\n\n"
            menu_text += "1️⃣  餐补/交补 (批量)\n"
            menu_text += "2️⃣  住宿补贴 (差额)\n"
            menu_text += "3️⃣  自定义补贴\n"
            menu_text += "4️⃣  预览记录\n"
            menu_text += "5️⃣  撤销上次操作\n"
            menu_text += "6️⃣  导出CSV\n"
            menu_text += "0️⃣  退出\n\n"
            menu_text += f"📊 当前: {len(self.rows)} 条记录"
            
            mode = simpledialog.askinteger("Moze报销生成器", menu_text, parent=self.root)
            if not mode or mode == 0: 
                break

            if mode == 1: # 批量模式
                d_in = simpledialog.askstring("步骤1/3 - 日期", 
                    "输入日期范围:\n\n"
                    "格式: yymmdd 或 yymmdd-yymmdd\n"
                    "示例: 250101 或 250101-250105")
                if not d_in: continue
                
                dates = self.parse_dates(d_in)
                if not dates: 
                    messagebox.showerror("格式错误", "日期格式应为 yymmdd\n示例: 250101 或 250101-250105")
                    continue
                
                loc = simpledialog.askstring("步骤2/3 - 地点", 
                    f"出差地点 (将写入描述字段):\n\n"
                    f"日期范围: {dates[0].strftime('%Y-%m-%d')} 至 {dates[-1].strftime('%Y-%m-%d')}\n"
                    f"共 {len(dates)} 天", 
                    initialvalue="武汉")
                if not loc: loc = ""
                
                # 询问补贴
                do_meal = messagebox.askyesno("步骤3/3 - 餐补", 
                    f"是否添加【餐费补贴】?\n\n"
                    f"标准: ¥{CFG['MEAL']}/天\n"
                    f"天数: {len(dates)}\n"
                    f"合计: ¥{CFG['MEAL'] * len(dates)}")
                
                do_trans = messagebox.askyesno("步骤3/3 - 交补", 
                    f"是否添加【交通补贴】?\n\n"
                    f"标准: ¥{CFG['TRANS']}/天\n"
                    f"天数: {len(dates)}\n"
                    f"合计: ¥{CFG['TRANS'] * len(dates)}")

                batch_count = 0
                for d in dates:
                    if do_meal: batch_count += self.add(d, "餐费补贴", CFG['MEAL'], loc, 0)
                    if do_trans: batch_count += self.add(d, "交通补贴", CFG['TRANS'], loc, 10)
                
                if batch_count > 0:
                    self.batch_sizes.append(batch_count)
                    total = (CFG['MEAL'] if do_meal else 0) + (CFG['TRANS'] if do_trans else 0)
                    messagebox.showinfo("添加成功", 
                        f"✅ 已添加 {len(dates)} 天记录\n\n"
                        f"餐补: {len(dates) if do_meal else 0} 天\n"
                        f"交补: {len(dates) if do_trans else 0} 天\n"
                        f"总计: ¥{total * len(dates)}")

            elif mode == 2: # 住宿模式
                d_in = simpledialog.askstring("住宿补贴 - 日期", 
                    "入住日期 (yymmdd):\n\n"
                    "示例: 250101")
                if not d_in: continue
                
                dates = self.parse_dates(d_in)
                if not dates: 
                    messagebox.showerror("格式错误", "日期格式应为 yymmdd")
                    continue
                
                loc = simpledialog.askstring("住宿补贴 - 地点", 
                    f"出差地点:\n\n"
                    f"日期: {dates[0].strftime('%Y-%m-%d')}", 
                    initialvalue="武汉")
                if not loc: loc = ""
                
                cost = simpledialog.askfloat("住宿补贴 - 金额", 
                    f"住宿补贴计算:\n\n"
                    f"公司标准: ¥{CFG['HOTEL']}\n"
                    f"请输入实际花费:")
                
                if cost is not None:
                    subsidy = CFG['HOTEL'] - cost
                    if subsidy > 0:
                        batch_count = self.add(dates[0], "住宿补贴", subsidy, loc, 120)
                        if batch_count > 0:
                            self.batch_sizes.append(batch_count)
                        messagebox.showinfo("添加成功", 
                            f"✅ 住宿补贴计算:\n\n"
                            f"标准: ¥{CFG['HOTEL']}\n"
                            f"实际: ¥{cost}\n"
                            f"补贴: ¥{subsidy}")
                    else:
                        messagebox.showwarning("无补贴", 
                            f"实际花费 (¥{cost}) {'超出' if subsidy < 0 else '等于'} 标准 (¥{CFG['HOTEL']})\n"
                            f"无需补贴")

            elif mode == 3: # 自定义补贴
                d_in = simpledialog.askstring("自定义补贴 - 日期", 
                    "日期 (yymmdd):\n\n"
                    "示例: 250101")
                if not d_in: continue
                
                dates = self.parse_dates(d_in)
                if not dates: continue
                
                name = simpledialog.askstring("自定义补贴 - 名称", 
                    "补贴名称:\n\n"
                    "示例: 加班补贴、其他补贴等")
                if not name: continue
                
                amt = simpledialog.askfloat("自定义补贴 - 金额", 
                    f"补贴金额 (元):\n\n"
                    f"名称: {name}\n"
                    f"日期: {dates[0].strftime('%Y-%m-%d')}")
                if not amt or amt <= 0: continue
                
                loc = simpledialog.askstring("自定义补贴 - 地点", 
                    "出差地点 (可选):", 
                    initialvalue="武汉")
                if not loc: loc = ""
                
                batch_count = self.add(dates[0], name, amt, loc, 60)
                if batch_count > 0:
                    self.batch_sizes.append(batch_count)
                messagebox.showinfo("添加成功", 
                    f"✅ 已添加自定义补贴:\n\n"
                    f"名称: {name}\n"
                    f"金额: ¥{amt}\n"
                    f"日期: {dates[0].strftime('%Y-%m-%d')}")

            elif mode == 4: # 预览
                self.show_preview()

            elif mode == 5: # 撤销
                self.undo_last_batch()

            elif mode == 6: # 导出
                if not self.rows: 
                    messagebox.showwarning("无记录", "没有可导出的记录")
                    continue
                
                # 保存逻辑
                path = Path(__file__).parent / "Moze4.0_Import"
                path.mkdir(exist_ok=True)
                timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
                f_name = path / f"Moze_报销_{timestamp}.csv"
                
                df = pd.DataFrame(self.rows, columns=COLUMNS)
                df.to_csv(f_name, index=False, encoding='utf-8-sig')
                
                total = df[df['记录类型'] == '收入']['金额'].sum()
                messagebox.showinfo("导出成功", 
                    f"✅ 文件已保存:\n\n"
                    f"📁 {f_name.name}\n"
                    f"📊 {len(self.rows)} 条记录\n"
                    f"💰 总金额: ¥{total:.2f}")
                break
                
        self.root.destroy()

if __name__ == "__main__":
    MozeReimburse().run()
