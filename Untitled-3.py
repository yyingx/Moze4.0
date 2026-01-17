# -*- coding: utf-8 -*-
"""
Moze 4.0 OCR 修复版 (优先匹配金额 + 调试模式)
"""
import os

# ========================================================
# 🚑 防报错补丁 (必须在最前)
# ========================================================
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import easyocr
import re
import datetime
import tkinter as tk
from tkinter import filedialog
import pandas as pd

# --- 配置 ---
DEFAULT_YEAR = 2026
MOZE_COLUMNS = ['账户', '币种', '记录类型', '主类别', '子类别', '金额', '手续费', '折扣', '名称', '商家', '日期', '时间', '项目', '描述', '标签', '对象']

# --- v11.63 分类字典 (保持不变) ---
DATA_SOURCE = {
    'MEAL': ["食堂", "餐厅", "餐饮", "美食", "面馆", "麦当劳", "肯德基", "星巴克", "茶", "咖啡", "烧烤", "麻辣烫", "火锅"],
    'SHOPPING': ["超市", "便利店", "百货", "商城", "优衣库", "淘宝", "京东", "拼多多", "唯品会"],
    'TRANSPORT': ["滴滴", "打车", "出行", "铁路", "火车", "地铁", "公交", "加油", "石化", "特来电", "充电"],
    'LIFE': ["电费", "水费", "燃气", "物业", "话费", "联通", "移动", "电信", "宽带"],
    'FRUIT': ["水果", "百果园", "鲜果"],
    'VEGETABLE': ["生鲜", "买菜", "叮咚", "盒马", "菜市场"]
}
CATEGORY_MAP = {
    'MEAL': ('支出', '饮食', '正餐'), 'SHOPPING': ('支出', '购物', '日用&家用'),
    'TRANSPORT': ('支出', '交通', '公共交通'), 'LIFE': ('支出', '居家', '水电煤'),
    'FRUIT': ('支出', '饮食', '饮料水果'), 'VEGETABLE': ('支出', '饮食', '食材')
}

class BColors:
    OKGREEN = '\033[92m'; WARNING = '\033[93m'; FAIL = '\033[91m'; ENDC = '\033[0m'

class MozeOCRTool:
    def __init__(self):
        print(f"{BColors.OKGREEN}🚀 初始化 OCR 引擎...{BColors.ENDC}")
        self.reader = easyocr.Reader(['ch_sim', 'en'], gpu=False) 
        self.rows = []

    def select_images(self):
        root = tk.Tk(); root.withdraw()
        return filedialog.askopenfilenames(title="请选择账单截图", filetypes=[("图片", "*.jpg *.png *.jpeg *.bmp")])

    def clean_amount(self, text):
        """ 金额提取逻辑 (修复常见OCR错误) """
        # 1. 修复符号: 把 "—", "_", "一" 在数字前修成 "-"
        text = re.sub(r'[_一—]\s*(\d)', r'-\1', text) 
        text = text.replace('Y', '').replace('¥', '').replace('￥', '')
        
        # 2. 严格匹配: 必须包含小数点 (如 12.00)
        match = re.search(r'([+\-]?)\s*([\d,]+\.\d{2})', text)
        if match:
            sign = match.group(1)
            num_str = match.group(2).replace(',', '')
            return sign, float(num_str)
        return None, None

    def parse_images(self, file_paths):
        # 日期正则: 01.15 或 1月15日
        date_pattern = re.compile(r'(\d{1,2})[月\.](\d{1,2})')
        current_date = datetime.datetime.now().strftime("%Y-%m-%d")
        
        print(f"\n{BColors.OKGREEN}>>> 开始识别...{BColors.ENDC}")

        for f_path in file_paths:
            print(f"📄 读取文件: {os.path.basename(f_path)}")
            try:
                with open(f_path, 'rb') as f: img_bytes = f.read()
                texts = self.reader.readtext(img_bytes, detail=0)
            except Exception as e:
                print(f"❌ 读取失败: {e}"); continue

            print(f"   (共 {len(texts)} 行文字，开始解析...)")
            
            # --- 核心循环 ---
            for i, text in enumerate(texts):
                text = text.strip()
                if not text: continue
                
                # [调试信息] 打印每一行看它是什么
                # print(f"   [DEBUG] 行{i}: {text}") 

                # 🛑 1. 先判断是不是【金额】(优先级最高!)
                sign, amount = self.clean_amount(text)
                if amount is not None:
                    # 找到金额！开始回溯找描述
                    desc = "OCR未识别"
                    # 往上找非日期、非金额的行作为描述
                    if i > 0:
                        prev = texts[i-1].strip()
                        if not date_pattern.search(prev) and self.clean_amount(prev)[1] is None:
                            desc = prev
                            # 再往上找一行拼凑 (针对长描述)
                            if i > 1:
                                prev_prev = texts[i-2].strip()
                                # 如果上上行很短且不是日期，可能是 "12:30" 这种时间，或者是描述的一部分
                                if len(prev_prev) > 2 and not date_pattern.search(prev_prev) and self.clean_amount(prev_prev)[1] is None:
                                    # 简单的逻辑：只有当它不是“支出/收入”这种字眼时才拼
                                    if "支出" not in prev_prev and "收入" not in prev_prev:
                                        desc = prev_prev + " " + desc

                    # 清洗描述
                    desc = desc.replace("转账", "").replace("转给", "").replace("来自", "").replace(":", "").strip()
                    
                    # 智能判断收支
                    rec_type = "支出" # 默认
                    if sign == '-' or "转给" in desc or "付款" in desc: rec_type = "支出"
                    elif sign == '+' or "收到" in desc or "退款" in desc: rec_type = "收入"

                    # 存入
                    self.rows.append({
                        '账户': '云闪付(OCR)', '币种': 'CNY', '记录类型': rec_type,
                        '主类别': '未分类', '子类别': '', '金额': abs(amount),
                        '手续费': 0, '折扣': 0, '名称': desc, '商家': '',
                        '日期': current_date, '时间': '00:00:00', '项目': '',
                        '描述': f"OCR源: {text}", '标签': '#OCR', '对象': ''
                    })
                    print(f"   ✅ [交易] {current_date} | {rec_type} {amount} | {desc}")
                    continue # 既然是金额，就处理完了，跳过后面的日期判断

                # 🛑 2. 再判断是不是【日期】
                date_match = date_pattern.search(text)
                if date_match:
                    try:
                        m, d = date_match.groups()
                        # 简单的逻辑：如果只匹配到 28.07 这种很可能是金额误判的，由于上面金额判断没过，这里可能会误判
                        # 增加一个补丁：如果这一行包含 "¥" 或者长得像数字，就别当日期了
                        if "¥" in text or len(re.findall(r'\d', text)) > 6:
                            pass 
                        else:
                            current_date = f"{DEFAULT_YEAR}-{int(m):02d}-{int(d):02d}"
                            print(f"   📅 [日期] 更新为: {current_date}")
                    except: pass
                    continue

    def save(self):
        if not self.rows:
            print(f"\n{BColors.FAIL}⚠️ 依然没有提取到数据！{BColors.ENDC}")
            print("👉 建议：请截图终端里输出的内容发给我，我来优化正则。")
            return

        df = pd.DataFrame(self.rows, columns=MOZE_COLUMNS)
        
        # 应用 v11.63 分类
        print("正在自动分类...")
        for key, keywords in DATA_SOURCE.items():
            rule = CATEGORY_MAP.get(key)
            if not rule: continue
            pattern = "|".join(keywords)
            mask = (df['名称'].str.contains(pattern, na=False)) & (df['主类别'] == '未分类')
            if mask.any():
                df.loc[mask, ['记录类型', '主类别', '子类别']] = rule
                
        # 转账特殊处理
        mask_tf = df['名称'].str.contains('应翔') & (df['记录类型']=='支出')
        if mask_tf.any():
            df.loc[mask_tf, ['主类别', '子类别']] = ['金融财务', '内部转账']

        ts = datetime.datetime.now().strftime('%H%M%S')
        filename = f"Moze_OCR_Fix_{ts}.csv"
        try:
            df.to_csv(filename, index=False, encoding='utf-8-sig')
            print(f"\n{BColors.OKGREEN}🎉 成功！文件已生成: {filename}{BColors.ENDC}")
            os.startfile(os.getcwd())
        except Exception as e: print(f"保存失败: {e}")

if __name__ == "__main__":
    app = MozeOCRTool()
    files = app.select_images()
    if files:
        app.parse_images(files)
        app.save()