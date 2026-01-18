# -*- coding: utf-8 -*-
"""
Moze 4.0 最终纯净版 (Production Ready)
版本: 2026.01.18_Clean
说明: 
1. 移除所有调试图片生成功能，只输出 CSV。
2. 包含所有已验证的核心逻辑 (智能拼接、倒序跨年、名字去噪、严格分类)。
"""

import os
# 🚑 防报错补丁
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import sys
import numpy as np
import cv2
import pandas as pd
import datetime
import re
import easyocr
import tkinter as tk
from tkinter import filedialog

# ==============================================================================
# ⚙️ 用户配置
# ==============================================================================
DEFAULT_YEAR = 2026  # 默认起始年份
MY_NAME = "应翔"
EXPORT_DIR = r"E:\天之逸2025\Moze4.0\Moze4.0_Import"

ACCOUNT_MAP = {
    '9708': '建设银行9780', '9780': '建设银行9780', '建设银行': '建设银行9780',
    '9579': '工商银行9579', '工商银行': '工商银行9579',
    '2973': '农业银行2973', '农业银行': '农业银行2973',
    '5680': '湖北省农村信用社', '农村信用社': '湖北省农村信用社', '信用社': '湖北省农村信用社',
    '4946': '平安银行4946', '平安银行': '平安银行4946',
    '1517': '兴业银行1517', '兴业银行': '兴业银行1517'
}

MOZE_COLUMNS = ['账户', '币种', '记录类型', '主类别', '子类别', '金额', '手续费',
                '折扣', '名称', '商家', '日期', '时间', '项目', '描述', '标签', '对象']

class BColors:
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'

# ==============================================================================
# 🔪 物理切割器 (纯净版)
# ==============================================================================
class ProjectionCutter:
    def cut_image(self, image_path):
        print(f"   🔍 读取图片...")
        try:
            img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), -1)
            if img is None: return []
            h, w = img.shape[:2]
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
            projection = np.sum(binary, axis=1)
            threshold = 255 * 3
            is_empty = projection < threshold
            cut_y_points = [0]; in_gap = False; start_gap = 0
            for y in range(h):
                if is_empty[y]:
                    if not in_gap: start_gap = y; in_gap = True
                else:
                    if in_gap:
                        end_gap = y
                        if end_gap - start_gap > 10:
                            center = (start_gap + end_gap) // 2
                            if center - cut_y_points[-1] > 20: cut_y_points.append(center)
                        in_gap = False
            if h - cut_y_points[-1] > 20: cut_y_points.append(h)
            slices = []
            for i in range(len(cut_y_points) - 1):
                y1 = cut_y_points[i]; y2 = cut_y_points[i+1]
                slices.append(img[y1:y2, 0:w]) # 不再需要 offset
            return slices
        except Exception as e:
            print(f"切割出错: {e}")
            return []

# ==============================================================================
# 🧠 核心解析器
# ==============================================================================
class MozeOCRTool:
    def __init__(self):
        print(f"{BColors.OKGREEN}🚀 Moze 4.0 纯净运行版启动...{BColors.ENDC}")
        self.reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)
        self.cutter = ProjectionCutter()
        self.rows = []

    def select_images(self):
        root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
        return filedialog.askopenfilenames(title="请选择账单截图")

    def clean_money(self, s):
        # 保留数字、点、逗号、正负号
        return re.sub(r'[^\d\.,\+\-]', '', s)

    def group_text_by_lines(self, results, y_threshold=15):
        if not results: return []
        boxes = []
        for r in results:
            box = r[0]; y_center = (box[0][1] + box[2][1]) / 2
            boxes.append({'y': y_center, 'data': r})
        boxes.sort(key=lambda x: x['y'])
        lines = []
        if not boxes: return []
        current_line = [boxes[0]['data']]; current_y = boxes[0]['y']
        for i in range(1, len(boxes)):
            item = boxes[i]
            if abs(item['y'] - current_y) < y_threshold:
                current_line.append(item['data'])
            else:
                current_line.sort(key=lambda r: r[0][0][0])
                lines.append(current_line)
                current_line = [item['data']]; current_y = item['y']
        if current_line:
            current_line.sort(key=lambda r: r[0][0][0])
            lines.append(current_line)
        return lines

    def parse_images(self, file_paths):
        date_pattern = re.compile(r'^(\d{1,2})[\.月\-/](\d{1,2})') 
        time_pattern = re.compile(r'(\d{1,2})[:：\.\s](\d{2})')

        print(f"\n{BColors.OKGREEN}>>> 开始处理 (无Debug图模式)...{BColors.ENDC}")

        for f_path in file_paths:
            print(f"\n📄 文件: {os.path.basename(f_path)}")
            
            # 智能年份状态机
            current_processing_year = DEFAULT_YEAR
            last_month = -1 
            
            slices = self.cutter.cut_image(f_path)
            if not slices: continue
            
            global_current_date = f"{current_processing_year}/01/01"

            for img_slice in slices:
                results = self.reader.readtext(img_slice)
                if not results: continue
                lines = self.group_text_by_lines(results)
                pending_top = None 

                for line_items in lines:
                    full_text = " ".join([r[1] for r in line_items])
                    
                    # 1. 银行行封杀令
                    is_bank_row = any(k in full_text for k in ["银行", "信用社", "钱包"])
                    
                    # 2. 日期行检测
                    dm = date_pattern.match(full_text)
                    has_money_sign = any(c in full_text for c in ['¥', '元', '+', '-'])
                    is_footer_year = re.search(r'202\d', full_text) and len(full_text) < 10
                    
                    if dm and not has_money_sign and len(full_text) < 15 and not is_bank_row and not is_footer_year:
                        m, d = map(int, dm.groups())
                        
                        # 倒序跨年逻辑
                        if last_month != -1:
                            if m > last_month: 
                                current_processing_year -= 1
                                print(f"   📉 [年份自动调整] -> {current_processing_year}")
                        last_month = m
                        global_current_date = f"{current_processing_year}/{m}/{d}"
                        print(f"   📅 [日期] {global_current_date}")
                        
                        if pending_top:
                            self.add_record_with_logic(pending_top, None, global_current_date)
                            pending_top = None
                        continue 

                    # 3. 交易行(Top)检测 - 智能拼接
                    has_amount = False
                    amount_val = 0.0
                    
                    if not is_bank_row:
                        # 倒序遍历
                        for i in range(len(line_items) - 1, -1, -1):
                            r = line_items[i]
                            txt = r[1]
                            clean = self.clean_money(txt)
                            
                            if '.' in clean and len(clean) > 2:
                                try:
                                    val = float(clean.replace(',', ''))
                                    # 过滤时间误判
                                    if (0 <= val <= 24.59) and not any(s in txt for s in ['+', '-', '¥', '元']):
                                        if not any(k in full_text for k in ["转", "收", "付", "支"]):
                                            continue 
                                except: continue
                                
                                # 智能拼接
                                merged_str = txt
                                j = i - 1
                                while j >= 0:
                                    left_r = line_items[j]
                                    left_txt = left_r[1]
                                    if re.match(r'^[\d,\.\+\-¥Y半\s]+$', left_txt):
                                        merged_str = left_txt + merged_str
                                        j -= 1
                                    else:
                                        break 
                                
                                final_clean = merged_str.replace('半', '').replace('Y', '').replace(' ', '').replace(',', '')
                                final_clean = re.sub(r'[^\d\.\-]', '', final_clean)
                                
                                try:
                                    amount_val = float(final_clean)
                                    has_amount = True
                                    break 
                                except: continue

                    has_keyword = any(k in full_text for k in ["转给", "收到", "付款", "来自", "支出"])
                    
                    if (has_amount or has_keyword) and not is_bank_row:
                        if pending_top:
                            self.add_record_with_logic(pending_top, None, global_current_date)
                            pending_top = None

                        target_name = ""
                        clean_text = full_text.replace(":", "").replace("：", "")
                        if "转给" in clean_text: target_name = clean_text.split("转给")[1]
                        elif "来自" in clean_text: target_name = clean_text.split("来自")[1]
                        if not target_name:
                            temp = full_text
                            for k in ["转账", "付款", "支出", "还款", "收到", "元", "¥"]:
                                temp = temp.replace(k, "")
                            target_name = temp
                        
                        # 名字去噪
                        target_name = target_name.replace("壬", "").replace("酃", "").replace("鬣", "").replace("半", "")
                        target_name = re.sub(r'[0-9\.,\-\+¥]', '', target_name).strip().split(' ')[0]

                        pending_top = {
                            "amount": amount_val,
                            "name": target_name,
                            "raw": full_text
                        }
                        continue

                    # 4. 信息行(Bottom)检测
                    has_time = time_pattern.search(full_text)
                    if (has_time or is_bank_row) and pending_top:
                        final_time = "00:00:00"
                        if has_time:
                            h, m_val = int(has_time.group(1)), int(has_time.group(2))
                            final_time = f"{h:02d}:{m_val:02d}:00"
                            
                        account_clean_text = full_text
                        if has_time: account_clean_text = account_clean_text.replace(has_time.group(0), "")
                        
                        source_account = "云闪付(未知)"
                        for k, v in ACCOUNT_MAP.items():
                            if k in account_clean_text: source_account = v; break
                        if source_account == "云闪付(未知)":
                            match = re.search(r'([\u4e00-\u9fa5]+银行|[\u4e00-\u9fa5]+信用社)', account_clean_text)
                            if match:
                                source_account = match.group(1)
                                num_match = re.search(r'\[(\d+)\]', account_clean_text)
                                if num_match: source_account += num_match.group(1)

                        bottom_data = {
                            "time": final_time,
                            "account": source_account,
                            "raw": full_text
                        }
                        self.add_record_with_logic(pending_top, bottom_data, global_current_date)
                        pending_top = None

                if pending_top:
                    self.add_record_with_logic(pending_top, None, global_current_date)
                    pending_top = None

    def add_record_with_logic(self, top, bottom, date_str):
        amount = abs(top["amount"])
        full_text = top["raw"]
        target_name = top["name"]
        
        time_str = "00:00:00"
        account = "云闪付(未知)"
        if bottom:
            time_str = bottom["time"]
            account = bottom["account"]

        final_type = ""; final_main = "未分类"; final_sub = ""; final_object = ""
        final_amount = 0.0

        if "转给" in full_text or "付款" in full_text or "支出" in full_text:
            if MY_NAME in target_name:
                final_type = "转出"; final_main = "转账"; final_sub = "转账"
                final_amount = -amount
                print(f"   🔄 [转出] {amount}")
            else:
                final_type = "应收款项"; final_main = "应收款项"; final_sub = "借出"
                final_amount = -amount
                final_object = target_name
                print(f"   📒 [借出] {final_object} {amount}")

        elif "来自" in full_text or "收到" in full_text:
            if MY_NAME in target_name:
                final_type = "转入"; final_main = "转账"; final_sub = "转账"
                final_amount = amount
                print(f"   🔄 [转入] {amount}")
            else:
                final_type = "应付款项"; final_main = "应付款项"; final_sub = "借入"
                final_amount = amount
                final_object = target_name
                print(f"   📒 [借入] {final_object} {amount}")
        
        else:
            if "还款" in full_text:
                final_type = "转出"; final_main = "财务"; final_sub = "信用卡还款"
                final_amount = -amount
                print(f"   💳 [还款] {amount}")
            else:
                final_type = "支出"
                final_amount = -amount
                print(f"   ✅ [支出] {amount}")

        row = {
            '账户': account, '币种': 'CNY', '记录类型': final_type,
            '主类别': final_main, '子类别': final_sub,
            '金额': final_amount,
            '手续费': 0, '折扣': 0, 
            '名称': '', 
            '商家': '', 
            '日期': date_str, '时间': time_str,
            '项目': '', '描述': '', '标签': '', '对象': final_object
        }
        self.rows.append(row)

    def save(self):
        if not self.rows: print("⚠️ 无数据"); return
        df = pd.DataFrame(self.rows, columns=MOZE_COLUMNS)
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"MOZE_Clean_{ts}.csv"
        if not os.path.exists(EXPORT_DIR): os.makedirs(EXPORT_DIR, exist_ok=True)
        full_path = os.path.join(EXPORT_DIR, filename)
        df.to_csv(full_path, index=False, encoding='utf-8-sig')
        print(f"\n{BColors.OKGREEN}🎉 导出成功: {filename}{BColors.ENDC}")
        os.startfile(os.path.dirname(full_path))

if __name__ == "__main__":
    app = MozeOCRTool()
    files = app.select_images()
    if files:
        app.parse_images(files)
        app.save()