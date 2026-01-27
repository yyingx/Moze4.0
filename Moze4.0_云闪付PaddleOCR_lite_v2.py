# -*- coding: utf-8 -*-
"""
Moze 4.0 审计日志版 (Audit Log) - PaddleOCR 稳定版 (终极过滤+无描述)
版本: 2026.01.27_v2_BugFix
环境要求: paddlepaddle==2.6.2, paddleocr==2.7.3
更新内容:
1. 过滤网速、月份汇总、收支统计行
2. 描述列默认为空 ("")
BUG修复 (v2):
1. 时间过滤阈值 24.59 -> 23.59
2. 裸 except 改为具体异常捕获
3. 时间修复逻辑优化，避免错配
"""

import os
import sys
import numpy as np
import cv2
import pandas as pd
import datetime
import re
import logging
import tkinter as tk
from tkinter import filedialog
from pathlib import Path

# === 引入 PaddleOCR ===
from paddleocr import PaddleOCR

# 屏蔽 Paddle 的繁杂日志
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
logging.getLogger("ppocr").setLevel(logging.WARNING)

# ==============================================================================
# ⚙️ 用户配置
# ==============================================================================
DEFAULT_YEAR = 2026
MY_NAME = "应翔"
CURRENT_DIR = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXPORT_DIR = CURRENT_DIR / "Moze4.0_Import"

# 🎯 数据过滤配置
EXCLUDE_RECEIVABLES = False

ACCOUNT_MAP = {
    '9708': '建设银行Ⅱ', '9579': '工商银行', '2973': '农业银行',
    '5680': '湖北农信', '4946': '平安银行4946', '8045': '平安银行', '1517': '兴业银行'
}

SPECIAL_ACCOUNT_MAP = {
    '活期+': '云闪付', '活期': '云闪付', '余额': '云闪付',
    '零钱': '云闪付', '云闪付钱包': '云闪付'
}

MOZE_COLUMNS = ['账户', '币种', '记录类型', '主类别', '子类别', '金额', '手续费',
                '折扣', '名称', '商家', '日期', '时间', '项目', '描述', '标签', '对象']

class BColors:
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'

# ==============================================================================
# 🔪 物理切割器
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
            is_empty = projection < (255 * 3)
            cut_y_points = [0]
            in_gap = False
            start_gap = 0
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
                y1 = cut_y_points[i]
                y2 = cut_y_points[i+1]
                slices.append(img[y1:y2, 0:w])
            return slices
        except (cv2.error, IOError, ValueError) as e:
            print(f"切割出错: {e}")
            return []

# ==============================================================================
# 🧠 核心解析器
# ==============================================================================
class MozeOCRTool:
    def __init__(self):
        print(f"{BColors.OKGREEN}🚀 Moze 4.0 终极版 v2 (BugFix) 启动...{BColors.ENDC}")
        try:
            self.ocr = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
        except Exception as e:
            print(f"   ⚠️ OCR初始化警告: {e}")
            self.ocr = PaddleOCR(use_angle_cls=True, lang="ch")
        self.cutter = ProjectionCutter()
        self.rows = []

    def clean_object_name(self, name):
        if not name: return name
        suffixes = ['转账', '付款', '收款', '借款', '还款', '车', '房', '费', '钱', '款', '支付', '收到', '来自']
        cleaned = name
        for s in suffixes:
            if cleaned.endswith(s): cleaned = cleaned[:-len(s)]
        prefixes = ['来自', '转给', '付给', '收到']
        for p in prefixes:
            if cleaned.startswith(p): cleaned = cleaned[len(p):]
        cleaned = re.sub(r'[^\u4e00-\u9fa5a-zA-Z]', '', cleaned).strip()
        return cleaned if cleaned else name

    def readtext(self, img):
        try:
            result = self.ocr.ocr(img, cls=True)
            if not result or result[0] is None: return []
            formatted = []
            for line in result[0]:
                if line[1][1] > 0.5: formatted.append([line[0], line[1][0], line[1][1]])
            return formatted
        except (AttributeError, IndexError, TypeError) as e:
            print(f"   ⚠️ OCR识别异常: {e}")
            return []

    def select_images(self):
        root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
        return filedialog.askopenfilenames(title="请选择账单截图")

    def clean_money(self, s):
        return re.sub(r'[^\d\.,\+\-_\—lI|\[\]]', '', s)

    def group_text_by_lines(self, results, y_threshold=15):
        if not results: return []
        boxes = [{'y': (r[0][0][1] + r[0][2][1])/2, 'data': r} for r in results]
        boxes.sort(key=lambda x: x['y'])
        lines = []
        if not boxes: return []
        curr_line = [boxes[0]['data']]
        curr_y = boxes[0]['y']
        for i in range(1, len(boxes)):
            item = boxes[i]
            if abs(item['y'] - curr_y) < y_threshold: curr_line.append(item['data'])
            else:
                curr_line.sort(key=lambda r: r[0][0][0])
                lines.append(curr_line)
                curr_line = [item['data']]
                curr_y = item['y']
        if curr_line: curr_line.sort(key=lambda r: r[0][0][0]); lines.append(curr_line)
        return lines

    def parse_images(self, file_paths):
        date_pattern = re.compile(r'^(\d{1,2})[\.月\-/](\d{1,2})')
        time_pattern = re.compile(r'\b(\d{1,2})[:：](\d{2})\b')

        print(f"\n{BColors.OKGREEN}>>> 开始处理...{BColors.ENDC}")

        for f_path in file_paths:
            print(f"\n📄 文件: {os.path.basename(f_path)}")
            curr_year = DEFAULT_YEAR
            last_month = -1
            slices = self.cutter.cut_image(f_path)
            if not slices: continue

            global_curr_date = f"{curr_year}/01/01"

            for img_slice in slices:
                results = self.readtext(img_slice)
                if not results: continue
                lines = self.group_text_by_lines(results)
                pending_top = None

                for line_items in lines:
                    full_text = " ".join([r[1] for r in line_items])

                    # =====================================================
                    # 🛡️ 核心过滤区
                    # =====================================================
                    if re.search(r'\d+[KkMm]?[Bb]?/s', full_text) or "/s" in full_text: continue
                    if "支出" in full_text and "收入" in full_text:
                        print(f"   🗑️ [过滤统计] {full_text}")
                        continue
                    if (("月" in full_text or "本月" in full_text) and
                        any(k in full_text for k in ["支出", "收入", "结余", "统计", "笔"]) and
                        "日" not in full_text):
                        print(f"   🗑️ [过滤汇总] {full_text}")
                        continue
                    # =====================================================

                    is_bank = any(k in full_text for k in ["银行", "信用社", "钱包"])
                    dm = date_pattern.match(full_text)
                    is_money = any(c in full_text for c in ['¥', '元', '+', '-'])

                    if dm and not is_money and len(full_text) < 15 and not is_bank:
                        m, d = map(int, dm.groups())
                        if last_month != -1 and m > last_month:
                            curr_year -= 1
                            print(f"   📉 [年份调整] -> {curr_year}")
                        last_month = m
                        global_curr_date = f"{curr_year}/{m}/{d}"
                        print(f"   📅 [日期] {global_curr_date}")
                        if pending_top: self.add_rec(pending_top, None, global_curr_date); pending_top = None
                        continue

                    has_amt = False
                    amt_val = 0.0
                    if not is_bank:
                        for i in range(len(line_items)-1, -1, -1):
                            txt = line_items[i][1]
                            clean = self.clean_money(txt)
                            if '.' in clean and len(clean) > 2:
                                try:
                                    val = float(clean.replace('l','1').replace('|','1'))
                                    # [BUG FIX] 时间阈值从 24.59 改为 23.59
                                    if (0 <= val <= 23.59) and not any(s in txt for s in ['+','-','¥']): continue
                                    amt_val = val; has_amt = True; break
                                except ValueError:
                                    pass

                    has_key = any(k in full_text for k in ["转给", "收到", "付款", "来自", "支出"])
                    if (has_amt or has_key) and not is_bank:
                        if pending_top: self.add_rec(pending_top, None, global_curr_date)

                        target = full_text
                        for k in ["转给","来自","收到"]:
                            if k in target: target = target.split(k)[1]

                        target = re.sub(r'[\d\.\,\-\+¥元]+', '', target)
                        target = self.clean_object_name(target)

                        pending_top = {"amount": amt_val, "name": target, "raw": full_text}
                        continue

                    tm = time_pattern.search(full_text)
                    if (tm or is_bank) and pending_top:
                        ft = "00:00:00"
                        if tm:
                            h, m_val = map(int, tm.groups())
                            # [BUG FIX] 处理更多 OCR 识别错误
                            if m_val == 71: m_val = 11
                            if m_val == 77: m_val = 17
                            if h<=23 and m_val<=59: ft = f"{h:02d}:{m_val:02d}:00"

                        acct = "云闪付(未知)"
                        for k,v in SPECIAL_ACCOUNT_MAP.items():
                            if k in full_text: acct = v; break
                        if acct == "云闪付(未知)":
                            for k,v in ACCOUNT_MAP.items():
                                if k in full_text: acct = v; break
                        if acct == "云闪付(未知)":
                            bm = re.search(r'([\u4e00-\u9fa5]+银行)', full_text)
                            if bm: acct = bm.group(1)

                        self.add_rec(pending_top, {"time": ft, "account": acct}, global_curr_date)
                        pending_top = None

                if pending_top: self.add_rec(pending_top, None, global_curr_date); pending_top = None

    def add_rec(self, top, bottom, date):
        amt = abs(top["amount"])
        name = top["name"]
        raw = top["raw"]
        time = bottom["time"] if bottom else "00:00:00"
        acct = bottom["account"] if bottom else "云闪付(未知)"

        rtype, main, sub, obj = "支出", "未分类", "", ""
        real = -amt

        if any(k in raw for k in ["转给", "付款", "支出", "还款"]):
            if MY_NAME in name: rtype="转出"; main="转账"; sub="转账"; real=-amt
            elif "还款" in raw: rtype="转出"; main="财务"; sub="信用卡还款"; real=-amt
            else: rtype="应收款项"; main="应收款项"; sub="借出"; real=-amt; obj=name
        elif any(k in raw for k in ["来自", "收到"]):
            if MY_NAME in name: rtype="转入"; main="转账"; sub="转账"; real=amt
            else: rtype="应付款项"; main="应付款项"; sub="借入"; real=amt; obj=name

        if EXCLUDE_RECEIVABLES and main in ['应收款项', '应付款项']: return

        print(f"   ✅ {date} {time} {main} {real} ({obj})")

        self.rows.append({
            '账户': acct, '币种': 'CNY', '记录类型': rtype,
            '主类别': main, '子类别': sub, '金额': real,
            '手续费': 0, '折扣': 0, '名称': '', '商家': '',
            '日期': date, '时间': time, '项目': '', '描述': '',
            '标签': '#UnionPay', '对象': obj
        })

    def save(self):
        if not self.rows: return print("⚠️ 无数据")
        df = pd.DataFrame(self.rows, columns=MOZE_COLUMNS)

        # [BUG FIX] 优化时间修复逻辑，增加更多匹配条件
        zeros = df[df['时间']=='00:00:00']
        for i, r in zeros.iterrows():
            # 同一天、同金额、同对象的记录
            m = df[(df['日期']==r['日期']) &
                   (abs(df['金额'])==abs(r['金额'])) &
                   (df['对象']==r['对象']) &
                   (df['时间']!='00:00:00')]
            if not m.empty:
                df.at[i, '时间'] = m.iloc[0]['时间']
                continue
            # 退而求其次：同一天、同金额
            m = df[(df['日期']==r['日期']) &
                   (abs(df['金额'])==abs(r['金额'])) &
                   (df['时间']!='00:00:00')]
            if not m.empty:
                df.at[i, '时间'] = m.iloc[0]['时间']

        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        fp = EXPORT_DIR / f"MOZE_UnionPay_{ts}.csv"
        os.makedirs(EXPORT_DIR, exist_ok=True)
        df.to_csv(fp, index=False, encoding='utf-8-sig')
        print(f"\n🎉 导出成功: {fp}"); os.startfile(EXPORT_DIR)

if __name__ == "__main__":
    app = MozeOCRTool()
    files = app.select_images()
    if files: app.parse_images(files); app.save()
