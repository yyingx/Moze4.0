# -*- coding: utf-8 -*-
"""
Moze 4.0 OCR 最终完美修正版
修复内容：
1. 修复时间识别 (00:00:00 -> 22:31:00)
2. 修复日期误判 (2026/22/31 -> 增加月份1-12校验)
3. 严格单行模式 (所见即所得，不生成双向记录)
4. 格式严格对齐云闪付标准 (YYYY/M/D)
"""

import os
import sys

# ==============================================================================
# 🚑 防报错补丁
# ==============================================================================
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import easyocr
import re
import datetime
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import copy
import time

# --- 核心配置 ---
DEFAULT_YEAR = 2026
MOZE_COLUMNS = ['账户', '币种', '记录类型', '主类别', '子类别', '金额', '手续费', '折扣', '名称', '商家', '日期', '时间', '项目', '描述', '标签', '对象']
EXPORT_DIR = r"E:\天之逸2025\Moze4.0\Moze4.0_Import"

# ==============================================================================
# ⚙️ 配置区
# ==============================================================================

# 1. 🏦 账户映射 (修正 OCR 识别的名字 -> 标准账户名)
ACCOUNT_MAP = {
    '9708': '建设银行9780', 
    '9780': '建设银行9780',
    '建设银行': '建设银行9780',
    '9579': '工商银行9579',
    '工商银行': '工商银行9579',
    '2973': '农业银行2973',
    '农业银行': '农业银行2973',
    '5680': '湖北省农村信用社',
    '农村信用社': '湖北省农村信用社',
    '信用社': '湖北省农村信用社'
}

# 2. 🔄 内部转账白名单 (仅用于标记，不再生成双向记录)
LIST_INTERNAL = ['应翔', '老婆', '老爸', '自己']

# ==============================================================================

class BColors:
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'

class MozeOCRTool:
    def __init__(self):
        print(f"{BColors.OKGREEN}🚀 程序启动中...{BColors.ENDC}")
        print("⏳ 正在加载 OCR 核心模型...", end="", flush=True)
        t0 = time.time()
        self.reader = easyocr.Reader(['ch_sim', 'en'], gpu=False) 
        print(f" {BColors.OKGREEN}✅ 加载成功 ({time.time()-t0:.1f}s){BColors.ENDC}")
        self.rows = []

    def select_images(self):
        root = tk.Tk(); root.withdraw()
        root.attributes('-topmost', True)
        return filedialog.askopenfilenames(title="请选择账单截图")

    def normalize_account(self, raw_account):
        if not raw_account: return "云闪付(待定)"
        for key, std_name in ACCOUNT_MAP.items():
            if key in raw_account:
                return std_name
        return raw_account

    def format_date(self, y, m, d):
        # 严格对齐标准：2026/1/15 (不补0)
        return f"{y}/{int(m)}/{int(d)}"

    def parse_images(self, file_paths):
        # 1. 日期正则 (匹配 1.15, 01.15, 1-15, 1月15)
        date_pattern = re.compile(r'^(\d{1,2})[\.月\-/](\d{1,2})') 
        
        # 2. 时间正则 (匹配 22:31, 22：31, 允许空格)
        time_pattern = re.compile(r'(\d{1,2})\s*[:：]\s*(\d{2})')
        
        # 3. 金额正则 (匹配 +428.07, -500, 一半100)
        amount_pattern = re.compile(r'([+\-一_半¥Y￥]*)\s*(\d{1,3}(?:,\d{3})*\.\d{2})')
        
        current_date_str = f"{DEFAULT_YEAR}/1/1"
        
        print(f"\n{BColors.OKGREEN}>>> 开始解析...{BColors.ENDC}")

        for f_path in file_paths:
            print(f"读取: {os.path.basename(f_path)}")
            try:
                with open(f_path, 'rb') as f: img_bytes = f.read()
                texts = self.reader.readtext(img_bytes, detail=0)
            except Exception as e:
                print(f"❌ 读取失败: {e}"); continue

            for i, text in enumerate(texts):
                text = text.strip()
                if not text: continue

                # [A] 锁定日期 (增加合法性校验，防止 22:31 误判为日期)
                # 只有长度较短，且月份1-12，日期1-31的才算日期
                if len(text) < 12 and re.match(r'^\d{1,2}[\.月\-/]\d{1,2}', text):
                    dm = date_pattern.search(text)
                    if dm:
                        m, d = map(int, dm.groups())
                        if 1 <= m <= 12 and 1 <= d <= 31:
                            current_date_str = self.format_date(DEFAULT_YEAR, m, d)
                            print(f"   📅 日期锁定: {current_date_str}")
                            continue # 是日期就跳过，不往下走

                # [B] 锁定金额 (作为核心锚点)
                am = amount_pattern.search(text)
                if am and not date_pattern.match(text):
                    amount_val = float(am.group(2).replace(',', ''))
                    
                    # === 1. 上下文分析 (找名字 & 方向) ===
                    prev_text = texts[i-1] if i > 0 else ""
                    prev_prev = texts[i-2] if i > 1 else ""
                    context = prev_prev + " " + prev_text
                    
                    is_out = False # 默认方向
                    
                    # 判断方向
                    if "转给" in context or "付款" in context: is_out = True
                    elif "来自" in context or "收到" in context: is_out = False
                    else:
                        # 符号辅助
                        raw_prefix = am.group(1)
                        if '-' in raw_prefix or '一半' in raw_prefix or '一' in raw_prefix: is_out = True
                        elif '+' in raw_prefix: is_out = False
                        else: is_out = True
                    
                    # 提取名字
                    clean_name = context.replace("转账", "").replace("转给", "").replace("来自", "").replace(":", "").replace("付款", "").replace("收到", "").strip()
                    clean_name = clean_name.split(" ")[-1] # 取最后一个词

                    # === 2. 向下分析 (找账户 & 时间 - 强力搜索) ===
                    raw_account = ""
                    time_str = "00:00:00"
                    
                    # 向下搜索 5 行
                    search_range = texts[i+1 : min(i+6, len(texts))]
                    
                    for line in search_range:
                        # 如果遇到新的金额行，立即停止，防止读到下一笔交易的时间
                        if amount_pattern.search(line) and not date_pattern.match(line):
                            break
                        
                        # 找账户
                        if not raw_account and ("银行" in line or "信用社" in line or "[" in line):
                            raw_account = re.split(r'[\[\(]', line)[0].strip()
                            num = re.search(r'(\d{4})', line)
                            if num: raw_account += num.group(1)
                        
                        # 找时间 (22:31)
                        tm = time_pattern.search(line)
                        if tm:
                            h, m_val = map(int, tm.groups())
                            # 校验时间合法性 (0-23, 0-59)
                            if 0 <= h <= 23 and 0 <= m_val <= 59:
                                time_str = f"{h:02d}:{m_val:02d}:00"

                    source_account = self.normalize_account(raw_account)

                    # === 3. 生成单行记录 ===
                    row = {
                        '账户': source_account, '币种': 'CNY', '记录类型': '', 
                        '主类别': '未分类', '子类别': '', '金额': 0,
                        '手续费': 0, '折扣': 0, '名称': clean_name,
                        '商家': '', '日期': current_date_str, '时间': time_str,
                        '项目': '', '描述': f"OCR源: {text}", '标签': '#OCR', '对象': ''
                    }

                    # 判断内部转账
                    is_internal = False
                    for int_name in LIST_INTERNAL:
                        if int_name in clean_name:
                            is_internal = True
                            break
                    
                    if is_internal:
                        # --- 场景1: 内部转账 (单行) ---
                        print(f"   🔄 内部转账: {source_account} ({'转出' if is_out else '转入'}) | 时间: {time_str}")
                        row['主类别'] = '转账'; row['子类别'] = '转账'
                        
                        if is_out:
                            row['记录类型'] = '转出'
                            row['金额'] = -abs(amount_val)
                        else:
                            row['记录类型'] = '转入'
                            row['金额'] = abs(amount_val)
                            
                        self.clean_fields(row, keep_object=None)
                        
                    elif clean_name and len(clean_name) > 1 and not re.search(r'\d', clean_name):
                        # --- 场景2: 借贷 (单行) ---
                        print(f"   📒 借贷往来: {clean_name} | 时间: {time_str}")
                        if is_out:
                            row['记录类型'] = '应收款项'
                            row['主类别'] = '应收款项'; row['子类别'] = '借出'
                            row['金额'] = -abs(amount_val)
                        else:
                            row['记录类型'] = '应付款项'
                            row['主类别'] = '应付款项'; row['子类别'] = '借入'
                            row['金额'] = abs(amount_val)
                        
                        self.clean_fields(row, keep_object=clean_name)

                    else:
                        # --- 场景3: 普通交易 ---
                        if is_out:
                            row['记录类型'] = '支出'; row['金额'] = -abs(amount_val)
                        else:
                            row['记录类型'] = '收入'; row['金额'] = abs(amount_val)
                        
                        print(f"   ✅ 普通交易: {row['金额']} | 时间: {time_str}")

                    self.rows.append(row)

    def clean_fields(self, row, keep_object=None):
        row['名称'] = ""; row['商家'] = ""; row['项目'] = ""
        row['描述'] = ""; row['标签'] = ""
        if keep_object: row['对象'] = keep_object
        else: row['对象'] = ""

    def save(self):
        if not self.rows:
            print(f"\n{BColors.WARNING}⚠️ 无数据{BColors.ENDC}"); return
        
        df = pd.DataFrame(self.rows, columns=MOZE_COLUMNS)
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"MOZE_UnionPay_{ts}.csv"
        
        if not os.path.exists(EXPORT_DIR):
            try: os.makedirs(EXPORT_DIR)
            except: pass
        
        full_path = os.path.join(EXPORT_DIR, filename) if os.path.exists(EXPORT_DIR) else filename

        try:
            df.to_csv(full_path, index=False, encoding='utf-8-sig')
            print(f"\n{BColors.OKGREEN}🎉 导出成功！{BColors.ENDC}")
            print(f"📄 文件: {filename}")
            print(f"📂 路径: {os.path.dirname(full_path)}")
            os.startfile(os.path.dirname(full_path))
        except Exception as e:
            print(f"❌ 保存失败: {e}")

if __name__ == "__main__":
    app = MozeOCRTool()
    files = app.select_images()
    if files:
        app.parse_images(files)
        app.save()