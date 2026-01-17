# -*- coding: utf-8 -*-
"""
Moze 4.0 OCR 最终极简实战版
核心逻辑 (简单粗暴)：
1. 名字匹配: 
   - 包含 'MY_NAME' (如: 应翔) -> 记为【转账】(转出/转入)
   - 包含 其他人名 -> 记为【借贷】(应收/应付 + 对象)
   - 没名字 -> 记为【收支】(支出/收入)
2. 图像处理: 自带增强，死磕时间识别 (22:31)
3. 字段清洗: 该空的空，该填的填
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
import time
from PIL import Image, ImageEnhance
import numpy as np

# --- 核心配置 ---
DEFAULT_YEAR = 2026
MOZE_COLUMNS = ['账户', '币种', '记录类型', '主类别', '子类别', '金额', '手续费', '折扣', '名称', '商家', '日期', '时间', '项目', '描述', '标签', '对象']
EXPORT_DIR = r"E:\天之逸2025\Moze4.0\Moze4.0_Import"

# ==============================================================================
# ⚙️ 账户与身份配置 (只改这里!)
# ==============================================================================

# 1. 🏦 账户映射
ACCOUNT_MAP = {
    '9708': '建设银行9780', '9780': '建设银行9780', '建设银行': '建设银行9780',
    '9579': '工商银行9579', '工商银行': '工商银行9579',
    '2973': '农业银行2973', '农业银行': '农业银行2973',
    '5680': '湖北省农村信用社', '农村信用社': '湖北省农村信用社', '信用社': '湖北省农村信用社'
}

# 2. 👤 您的名字 (账单上出现这个名字，就等于"转账")
# 别填"自己"这种虚词了，直接填账单上显示的那个字！
MY_NAME = "应翔"

# ==============================================================================

class BColors:
    OKGREEN = '\033[92m'; WARNING = '\033[93m'; FAIL = '\033[91m'; ENDC = '\033[0m'

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
            if key in raw_account: return std_name
        return raw_account

    def format_date(self, y, m, d):
        return f"{y}/{int(m)}/{int(d)}"

    def preprocess_image(self, file_path):
        """ 🎨 图像增强 (黑科技: 解决时间看不清 + OpenCV报错) """
        try:
            img = Image.open(file_path).convert('L')
            img = ImageEnhance.Contrast(img).enhance(2.5) # 对比度
            img = ImageEnhance.Sharpness(img).enhance(2.0) # 锐度
            # 二值化后立即转回L模式
            img = img.point(lambda x: 0 if x < 140 else 255, '1').convert('L')
            return np.array(img)
        except Exception as e:
            print(f"⚠️ 增强失败，使用原图: {e}")
            with open(file_path, 'rb') as f: return f.read()

    def parse_images(self, file_paths):
        # 正则
        date_pattern = re.compile(r'^(\d{1,2})[\.月\-/](\d{1,2})') 
        time_pattern = re.compile(r'(?:(\d{1,2})\s*[:：\.]\s*(\d{2}))|^(\d{2})(\d{2})$')
        amount_pattern = re.compile(r'([+\-一_半¥Y￥]*)\s*(\d{1,3}(?:,\d{3})*\.\d{2})')
        
        current_date_str = f"{DEFAULT_YEAR}/1/1"
        buffer_time = None
        
        print(f"\n{BColors.OKGREEN}>>> 开始解析...{BColors.ENDC}")

        for f_path in file_paths:
            print(f"读取: {os.path.basename(f_path)}")
            img_data = self.preprocess_image(f_path)
            
            try: texts = self.reader.readtext(img_data, detail=0)
            except Exception as e: print(f"❌ 读取失败: {e}"); continue

            for i, text in enumerate(texts):
                text = text.strip()
                if not text: continue

                # --- 1. 时间拦截器 (看到冒号就扣下) ---
                tm_check = time_pattern.search(text)
                if tm_check and not date_pattern.match(text) and not amount_pattern.search(text) and len(text) < 10:
                    h, m_val = 0, 0
                    if tm_check.group(1): h, m_val = int(tm_check.group(1)), int(tm_check.group(2))
                    elif tm_check.group(3) and text.isdigit(): h, m_val = int(tm_check.group(3)), int(tm_check.group(4))
                    if 0 <= h <= 23 and 0 <= m_val <= 59:
                        buffer_time = f"{h:02d}:{m_val:02d}:00"
                        continue 

                # --- 2. 日期锁定 ---
                match_date = False
                if len(text) < 20 and re.match(r'^\d{1,2}[\.月\-/]\d{1,2}', text):
                    dm = date_pattern.search(text)
                    if dm:
                        m, d = map(int, dm.groups())
                        if 1 <= m <= 12 and 1 <= d <= 31:
                            current_date_str = self.format_date(DEFAULT_YEAR, m, d)
                            print(f"   📅 日期锁定: {current_date_str}")
                            match_date = True
                            continue

                # --- 3. 金额锚点 (开始干活) ---
                am = amount_pattern.search(text)
                if am and not match_date:
                    amount_val = float(am.group(2).replace(',', ''))
                    
                    # 上下文 & 方向
                    prev_text = texts[i-1] if i > 0 else ""
                    prev_prev = texts[i-2] if i > 1 else ""
                    context = prev_prev + " " + prev_text
                    
                    is_out = False 
                    if "转给" in context or "付款" in context: is_out = True
                    elif "来自" in context or "收到" in context: is_out = False
                    else:
                        raw_prefix = am.group(1)
                        is_out = ('-' in raw_prefix or '一半' in raw_prefix or '一' in raw_prefix)

                    # 提取名字
                    clean_name = context.replace("转账", "").replace("转给", "").replace("来自", "").replace(":", "").replace("付款", "").replace("收到", "").strip()
                    clean_name = clean_name.split(" ")[-1]
                    if re.search(r'\d', clean_name): clean_name = "" # 名字里不能有数字

                    # 找账户 & 时间
                    raw_account = ""; final_time = buffer_time if buffer_time else "00:00:00"; buffer_time = None
                    search_range = texts[i+1 : min(i+8, len(texts))]
                    for line in search_range:
                        if amount_pattern.search(line) and not date_pattern.match(line): break
                        
                        if not raw_account and ("银行" in line or "信用社" in line or "[" in line):
                            raw_account = re.split(r'[\[\(]', line)[0].strip()
                            num = re.search(r'(\d{4})', line)
                            if num: raw_account += num.group(1)
                        
                        if final_time == "00:00:00":
                            tm = time_pattern.search(line)
                            if tm:
                                h, m_val = 0, 0
                                if tm.group(1): h, m_val = int(tm.group(1)), int(tm.group(2))
                                elif tm.group(3) and line.isdigit(): h, m_val = int(tm.group(3)), int(tm.group(4))
                                if 0 <= h <= 23 and 0 <= m_val <= 59: final_time = f"{h:02d}:{m_val:02d}:00"

                    source_account = self.normalize_account(raw_account)

                    # === 🔥 核心判决逻辑 ===
                    final_type = ""
                    final_main = "未分类"; final_sub = ""
                    final_name = ""; final_object = ""
                    
                    # 判决 1: 名字包含 "应翔" (MY_NAME) -> 转账
                    if MY_NAME in clean_name:
                        final_type = '转出' if is_out else '转入'
                        final_main = '转账'; final_sub = '转账'
                        print(f"   🔄 内部转账: {source_account} ({final_type})")
                        
                    # 判决 2: 名字是别人 -> 借贷
                    elif clean_name and len(clean_name) > 1:
                        if is_out:
                            final_type = '应收款项'; final_main = '应收款项'; final_sub = '借出'
                        else:
                            final_type = '应付款项'; final_main = '应付款项'; final_sub = '借入'
                        final_object = clean_name # 填对象
                        print(f"   📒 借贷往来: {final_object}")
                        
                    # 判决 3: 没名字 -> 普通收支
                    else:
                        final_type = '支出' if is_out else '收入'
                        final_name = clean_name
                        print(f"   ✅ 普通交易: {amount_val}")

                    # 生成行
                    row = {
                        '账户': source_account, '币种': 'CNY', '记录类型': final_type, 
                        '主类别': final_main, '子类别': final_sub, 
                        '金额': -abs(amount_val) if is_out else abs(amount_val),
                        '手续费': 0, '折扣': 0, 
                        '名称': final_name, # 转账/借贷时这里自动为空
                        '商家': '', '日期': current_date_str, '时间': final_time,
                        '项目': '', '描述': '', '标签': '', 
                        '对象': final_object # 只有借贷时这里才有值
                    }
                    self.rows.append(row)

    def save(self):
        if not self.rows: print(f"\n{BColors.WARNING}⚠️ 无数据{BColors.ENDC}"); return
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
        except Exception as e: print(f"❌ 保存失败: {e}")

if __name__ == "__main__":
    app = MozeOCRTool()
    files = app.select_images()
    if files:
        app.parse_images(files)
        app.save()