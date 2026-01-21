# -*- coding: utf-8 -*-
"""
Moze 4.0 审计日志版 (Audit Log)
版本: 2026.01.18_Audit
特性:
1. 【修正留痕】: 凡是程序自动修改了时间的，都会打印 "🔧 [时间同步]" 日志，显示修改前后的值。
2. 【透明汇报】: 统计最终修正了多少条时间数据。
3. 【全功能集成】: 包含智能拼接、倒序跨年、名字去噪、闭环验证等所有核心科技。
"""
import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
import sys  # noqa: E402
import numpy as np  # noqa: E402
import cv2  # noqa: E402
import pandas as pd  # noqa: E402
import datetime  # noqa: E402
import re  # noqa: E402
import easyocr  # noqa: E402
import tkinter as tk  # noqa: E402
from tkinter import filedialog  # noqa: E402
from pathlib import Path  # noqa: E402

# ==============================================================================
# ⚙️ 用户配置
# ==============================================================================
DEFAULT_YEAR = 2026
MY_NAME = "应翔"
CURRENT_DIR = Path(__file__).parent if '__file__' in locals() else Path.cwd()
EXPORT_DIR = CURRENT_DIR / "Moze4.0_Import"


# 🎯 数据过滤配置
# 设置为True时，将排除应收应付记录（借入/借出），只保留收入/支出/转账
EXCLUDE_RECEIVABLES = False  # 改为True可排除应收应付记录

# ⏰ 时间处理配置
# 设置为True时，会自动删除时间为00:00:00的记录（通常是OCR识别失败的记录）
EXCLUDE_ZERO_TIME = False  # 改为True可排除无时间信息的记录

# 🎯 智能过滤配置（推荐开启）
# 自动过滤可能是汇总数据的记录（时间00:00:00 + 账户未知 + 无对象名）
AUTO_FILTER_SUMMARY = True  # 建议保持True，自动识别并过滤汇总数据



ACCOUNT_MAP = {
    '9708': '建设银行Ⅱ',
    '9579': '工商银行',
    '2973': '农业银行',
    '5680': '湖北农信',
    '4946': '平安银行4946',
    '8045': '平安银行',
    '1517': '兴业银行1517'
}

# 🏦 特殊账户名映射（云闪付的虚拟账户）
SPECIAL_ACCOUNT_MAP = {
    '活期+': '云闪付',
    '活期': '云闪付',
    '余额': '云闪付',
    '零钱': '云闪付',
    '云闪付钱包': '云闪付'
}


MOZE_COLUMNS = ['账户', '币种', '记录类型', '主类别', '子类别', '金额', '手续费',
                '折扣', '名称', '商家', '日期', '时间', '项目', '描述', '标签', '对象']


class BColors:
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BLUE = '\033[94m'

# ==============================================================================
# 🔪 物理切割器
# ==============================================================================


class ProjectionCutter:
    def cut_image(self, image_path):
        print(f"   🔍 读取图片...")
        try:
            img = cv2.imdecode(np.fromfile(image_path, dtype=np.uint8), -1)
            if img is None:
                return []
            h, w = img.shape[:2]
            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            _, binary = cv2.threshold(
                gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
            projection = np.sum(binary, axis=1)
            threshold = 255 * 3
            is_empty = projection < threshold
            cut_y_points = [0]
            in_gap = False
            start_gap = 0
            for y in range(h):
                if is_empty[y]:
                    if not in_gap:
                        start_gap = y
                        in_gap = True
                else:
                    if in_gap:
                        end_gap = y
                        if end_gap - start_gap > 10:
                            center = (start_gap + end_gap) // 2
                            if center - cut_y_points[-1] > 20:
                                cut_y_points.append(center)
                        in_gap = False
            if h - cut_y_points[-1] > 20:
                cut_y_points.append(h)
            slices = []
            for i in range(len(cut_y_points) - 1):
                y1 = cut_y_points[i]
                y2 = cut_y_points[i+1]
                slices.append(img[y1:y2, 0:w])
            return slices
        except Exception as e:
            print(f"切割出错: {e}")
            return []

# ==============================================================================
# 🧠 核心解析器
# ==============================================================================


class MozeOCRTool:
    def __init__(self):
        print(f"{BColors.OKGREEN}🚀 Moze 4.0 审计日志版启动...{BColors.ENDC}")
        self.reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)
        self.cutter = ProjectionCutter()
        self.rows = []

    def select_images(self):
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        return filedialog.askopenfilenames(title="请选择账单截图")

    def clean_money(self, s):
        # 允许 l, I, | 等类似 1 的字符存在
        return re.sub(r'[^\d\.,\+\-_\—lI|\[\]]', '', s)

    def group_text_by_lines(self, results, y_threshold=15):
        if not results:
            return []
        boxes = []
        for r in results:
            box = r[0]
            y_center = (box[0][1] + box[2][1]) / 2
            boxes.append({'y': y_center, 'data': r})
        boxes.sort(key=lambda x: x['y'])
        lines = []
        if not boxes:
            return []
        current_line = [boxes[0]['data']]
        current_y = boxes[0]['y']
        for i in range(1, len(boxes)):
            item = boxes[i]
            if abs(item['y'] - current_y) < y_threshold:
                current_line.append(item['data'])
            else:
                current_line.sort(key=lambda r: r[0][0][0])
                lines.append(current_line)
                current_line = [item['data']]
                current_y = item['y']
        if current_line:
            current_line.sort(key=lambda r: r[0][0][0])
            lines.append(current_line)
        return lines

    def parse_images(self, file_paths):
        date_pattern = re.compile(r'^(\d{1,2})[\.月\-/](\d{1,2})')
        # 增强时间正则
        time_pattern = re.compile(r'(\d{1,2})[:：\.\s]+(\d{2})')

        print(f"\n{BColors.OKGREEN}>>> 开始处理...{BColors.ENDC}")

        for f_path in file_paths:
            print(f"\n📄 文件: {os.path.basename(f_path)}")

            current_processing_year = DEFAULT_YEAR
            last_month = -1

            slices = self.cutter.cut_image(f_path)
            if not slices:
                continue

            global_current_date = f"{current_processing_year}/01/01"

            for img_slice in slices:
                results = self.reader.readtext(img_slice)
                if not results:
                    continue
                lines = self.group_text_by_lines(results)
                pending_top = None

                for line_items in lines:
                    full_text = " ".join([r[1] for r in line_items])

                    is_bank_row = any(k in full_text for k in [
                                      "银行", "信用社", "钱包"])
                    dm = date_pattern.match(full_text)
                    has_money_sign = any(c in full_text for c in [
                                         '¥', '元', '+', '-'])
                    is_footer_year = re.search(
                        r'202\d', full_text) and len(full_text) < 10

                    if dm and not has_money_sign and len(full_text) < 15 and not is_bank_row and not is_footer_year:
                        m, d = map(int, dm.groups())
                        if last_month != -1:
                            if m > last_month:
                                current_processing_year -= 1
                                print(
                                    f"   📉 [年份自动调整] -> {current_processing_year}")
                        last_month = m
                        global_current_date = f"{current_processing_year}/{m}/{d}"
                        print(f"   📅 [日期] {global_current_date}")

                        if pending_top:
                            self.add_record_with_logic(
                                pending_top, None, global_current_date)
                            pending_top = None
                        continue

                    has_amount = False
                    amount_val = 0.0

                    if not is_bank_row:
                        for i in range(len(line_items) - 1, -1, -1):
                            r = line_items[i]
                            txt = r[1]
                            clean = self.clean_money(txt)

                            if '.' in clean and len(clean) > 2:
                                try:
                                    val = float(clean.replace(',', '').replace(
                                        'l', '1').replace('|', '1'))
                                    if (0 <= val <= 24.59) and not any(s in txt for s in ['+', '-', '¥', '元']):
                                        if not any(k in full_text for k in ["转", "收", "付", "支"]):
                                            continue
                                except:
                                    pass

                                merged_str = txt
                                j = i - 1
                                while j >= 0:
                                    left_r = line_items[j]
                                    left_txt = left_r[1]
                                    if re.match(r'^[\d,\.\+\-_\—¥Y半\s]+$', left_txt):
                                        merged_str = left_txt + merged_str
                                        j -= 1
                                    elif re.match(r'^[lI|\[\]]+$', left_txt):
                                        merged_str = "1" + merged_str
                                        j -= 1
                                    else:
                                        break

                                final_clean = merged_str.replace('半', '').replace(
                                    'Y', '').replace(' ', '').replace(',', '')
                                final_clean = final_clean.replace(
                                    '_', '-').replace('—', '-')
                                final_clean = final_clean.replace('l', '1').replace(
                                    'I', '1').replace('|', '1').replace('[', '1').replace(']', '1')
                                final_clean = re.sub(
                                    r'[^\d\.\-]', '', final_clean)

                                try:
                                    amount_val = float(final_clean)
                                    has_amount = True
                                    break
                                except:
                                    continue

                    has_keyword = any(k in full_text for k in [
                                      "转给", "收到", "付款", "来自", "支出"])

                    if (has_amount or has_keyword) and not is_bank_row:
                        if pending_top:
                            self.add_record_with_logic(
                                pending_top, None, global_current_date)
                            pending_top = None

                        target_name = ""
                        clean_text = full_text.replace(
                            ":", "").replace("：", "")
                        if "转给" in clean_text:
                            target_name = clean_text.split("转给")[1]
                        elif "来自" in clean_text:
                            target_name = clean_text.split("来自")[1]
                        elif "收到" in clean_text:
                            target_name = clean_text.split("收到")[1]
                        else:
                            temp = full_text
                            for k in ["转账", "付款", "支出", "还款", "收到", "元", "¥"]:
                                temp = temp.replace(k, "")
                            target_name = temp

                        target_name = target_name.replace("半", "").replace("收入", "").replace(
                            "支出", "").replace("一", "").replace("Y", "").replace("¥", "")
                        target_name = target_name.replace("壬", "王").replace(
                            "玍", "王").replace("酃", "").replace("鬣", "")
                        target_name = re.sub(
                            r'[^\u4e00-\u9fa5a-zA-Z]', '', target_name).strip()

                        pending_top = {
                            "amount": amount_val,
                            "name": target_name,
                            "raw": full_text
                        }
                        continue

                    has_time = time_pattern.search(full_text)
                    if (has_time or is_bank_row) and pending_top:
                        final_time = "00:00:00"
                        if has_time:
                            # 🔍 调试：打印OCR识别到的原始时间字符串
                            time_str_raw = has_time.group(0)
                            h_str = has_time.group(1)
                            m_str = has_time.group(2)
                            
                            print(f"      🔍 [OCR时间] 原始: '{time_str_raw}' -> 小时:{h_str} 分钟:{m_str}")
                            
                            h, m_val = int(h_str), int(m_str)
                            
                            # 🔧 时间验证与智能修正
                            original_time = f"{h:02d}:{m_val:02d}:00"  # 记录原始时间（修正前）
                            
                            # ⚠️ 检查明显的时间错误（可能是OCR识别顺序问题）
                            if h < 10 and m_val > 10:
                                # 如果小时是个位数，但分钟是20+，可能是识别反了
                                # 例如: 02:22 可能本来是 22:31 的错误识别
                                if m_val in [20, 22, 16]:  # 常见的错误模式
                                    print(f"      ⚠️ [可疑时间] {h:02d}:{m_val:02d} 看起来不太对，请检查原图")
                            
                            if m_val > 59:
                                # 可能是OCR错误：71可能是11，51可能是31等
                                if m_val == 71:
                                    m_val = 11  # 常见错误：7和1靠太近
                                    print(f"      🔧 [时间修正] OCR错误 {h}:71 -> {h}:11")
                                else:
                                    print(f"      ⚠️ [时间异常] {h}:{m_val} -> 00:00")
                                    h, m_val = 0, 0
                            
                            if h > 23:
                                print(f"      ⚠️ [时间异常] {h}:{m_val} -> 00:00")
                                h, m_val = 0, 0
                            
                            final_time = f"{h:02d}:{m_val:02d}:00"


                        account_clean_text = full_text
                        if has_time:
                            account_clean_text = account_clean_text.replace(
                                has_time.group(0), "")

                        source_account = "云闪付(未知)"
                        
                        # 🏦 优先检查特殊账户名（如"活期+"）
                        for special_name, moze_name in SPECIAL_ACCOUNT_MAP.items():
                            if special_name in account_clean_text:
                                source_account = moze_name
                                break
                        
                        # 如果没匹配到特殊账户，再检查银行卡后四位
                        if source_account == "云闪付(未知)":
                            for k, v in ACCOUNT_MAP.items():
                                if k in account_clean_text:
                                    source_account = v
                                    break
                        if source_account == "云闪付(未知)":
                            match = re.search(
                                r'([\u4e00-\u9fa5]+银行|[\u4e00-\u9fa5]+信用社)', account_clean_text)
                            if match:
                                source_account = match.group(1)
                                num_match = re.search(
                                    r'\[(\d+)\]', account_clean_text)
                                if num_match:
                                    source_account += num_match.group(1)

                        bottom_data = {
                            "time": final_time,
                            "original_time": original_time if "original_time" in locals() else final_time,
                            "account": source_account,
                            "raw": full_text
                        }
                        self.add_record_with_logic(
                            pending_top, bottom_data, global_current_date)
                        pending_top = None

                if pending_top:
                    self.add_record_with_logic(
                        pending_top, None, global_current_date)
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

        final_type = ""
        final_main = "未分类"
        final_sub = ""
        final_object = ""
        final_amount = 0.0

        if "转给" in full_text or "付款" in full_text or "支出" in full_text:
            if MY_NAME in target_name:
                final_type = "转出"
                final_main = "转账"
                final_sub = "转账"
                final_amount = -amount
            else:
                final_type = "应收款项"
                final_main = "应收款项"
                final_sub = "借出"
                final_amount = -amount
                final_object = target_name

        elif "来自" in full_text or "收到" in full_text:
            if MY_NAME in target_name:
                final_type = "转入"
                final_main = "转账"
                final_sub = "转账"
                final_amount = amount
            else:
                final_type = "应付款项"
                final_main = "应付款项"
                final_sub = "借入"
                final_amount = amount
                final_object = target_name

        else:
            if "还款" in full_text:
                final_type = "转出"
                final_main = "财务"
                final_sub = "信用卡还款"
                final_amount = -amount
            else:
                final_type = "支出"
                final_amount = -amount

        if final_type == "转出" and final_main == "转账":
            print(f"   🔄 [转出] {amount}")
        elif final_type == "转入":
            print(f"   🔄 [转入] {amount}")
        elif "借" in final_type or "借" in final_sub:
            print(f"   📒 [{final_sub}] {target_name} {amount}")
        else:
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
        # 🎯 根据配置过滤应收应付记录
        if EXCLUDE_RECEIVABLES and final_main in ['应收款项', '应付款项']:
            print(f"   🚫 [已过滤] {final_sub}: {target_name} {amount}")
            return
        
        self.rows.append(row)

    def verify_loops(self):
        print(f"\n{BColors.OKGREEN}>>> 正在进行智能对账 (闭环验证)...{BColors.ENDC}")
        df = pd.DataFrame(self.rows)
        if df.empty:
            return

        transfers = df[df['主类别'] == '转账'].copy()
        total_transfers = len(transfers)

        if total_transfers == 0:
            print("   ℹ️ 本次未检测到内部转账记录。")
            return

        matched_indices = set()
        fail_list = []
        success_pairs = 0
        fixed_time_count = 0

        for idx, row in transfers.iterrows():
            if idx in matched_indices:
                continue

            current_amount = row['金额']
            current_date = row['日期']

            potential_partners = transfers[
                (transfers['金额'] == -current_amount) &
                (transfers['日期'] == current_date) &
                (~transfers.index.isin(matched_indices)) &
                (transfers.index != idx)
            ]

            if not potential_partners.empty:
                partner_idx = potential_partners.index[0]
                partner_row = potential_partners.iloc[0]

                # === 🔥 时间同步与日志 🔥 ===
                t1 = row['时间']
                t2 = partner_row['时间']
                best_time = t1
                if t1 == "00:00:00" and t2 != "00:00:00":
                    best_time = t2
                elif t2 == "00:00:00" and t1 != "00:00:00":
                    best_time = t1

                # 检查并修正
                if self.rows[idx]['时间'] != best_time:
                    print(
                        f"      🔧 [时间同步] {current_date} ({current_amount}): {self.rows[idx]['时间']} -> {best_time}")
                    self.rows[idx]['时间'] = best_time
                    fixed_time_count += 1

                if self.rows[partner_idx]['时间'] != best_time:
                    print(
                        f"      🔧 [时间同步] {current_date} ({-current_amount}): {self.rows[partner_idx]['时间']} -> {best_time}")
                    self.rows[partner_idx]['时间'] = best_time
                    fixed_time_count += 1

                matched_indices.add(idx)
                matched_indices.add(partner_idx)
                success_pairs += 1
            else:
                fail_list.append(
                    f"{current_date} {row['时间']} ({current_amount})")

        # === 📊 最终汇报 ===
        print(f"   📊 统计: 共检测到 {total_transfers} 笔转账记录。")
        if fixed_time_count > 0:
            print(f"   🔧 修正: 共自动修复了 {fixed_time_count} 处时间错误。")

        if not fail_list:
            print(
                f"   🎉 {BColors.OKGREEN}完美！{success_pairs} 对转账全部闭环成功。{BColors.ENDC}")
        else:
            print(
                f"   ⚠️ {BColors.WARNING}警告: 发现 {len(fail_list)} 笔孤立转账，请核对:{BColors.ENDC}")
            for msg in fail_list:
                print(f"      - {msg}")


    def check_zero_times(self):
        """检查并警告00:00:00的记录"""
        print(f"\n{BColors.OKGREEN}>>> 检查时间完整性...{BColors.ENDC}")
        zero_time_records = [r for r in self.rows if r['时间'] == '00:00:00']
        
        if not zero_time_records:
            print("   ✅ 所有记录都包含时间信息。")
            return
        
        print(f"   ⚠️ {BColors.WARNING}发现 {len(zero_time_records)} 条记录缺少时间信息:{BColors.ENDC}")
        
        for record in zero_time_records:
            date = record['日期']
            amount = record['金额']
            account = record['账户']
            category = record['主类别']
            
            # 尝试从同一天的其他记录推测时间
            same_day_records = [r for r in self.rows 
                              if r['日期'] == date 
                              and r['时间'] != '00:00:00'
                              and abs(r['金额']) == abs(amount)]
            
            if same_day_records:
                suggested_time = same_day_records[0]['时间']
                print(f"      📌 {date} {category} {amount} ({account})")
                print(f"         💡 建议时间: {suggested_time} (来自同日同金额记录)")
            else:
                print(f"      📌 {date} {category} {amount} ({account})")
                print(f"         ⚠️  建议手动检查原始截图补充时间")

    def save(self):
        if not self.rows:
            print("⚠️ 无数据")
            return

        self.verify_loops()
        self.check_zero_times()

        df = pd.DataFrame(self.rows, columns=MOZE_COLUMNS)
        ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"MOZE_UnionPay_{ts}.csv"
        if not os.path.exists(EXPORT_DIR):
            os.makedirs(EXPORT_DIR, exist_ok=True)
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
