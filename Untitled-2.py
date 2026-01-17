# -*- coding: utf-8 -*-
import os

# ==============================================================================
# 🛑 第一步：防报错补丁 (必须放在最前面)
# ==============================================================================
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import easyocr
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import re
import datetime
import sys

# --- 核心配置 ---
DEFAULT_YEAR = 2026   # 默认年份 (因为截图里通常没有年份)
MOZE_COLUMNS = ['账户', '币种', '记录类型', '主类别', '子类别', '金额', '手续费', '折扣', '名称', '商家', '日期', '时间', '项目', '描述', '标签', '对象']

class MozeOCRTool:
    def __init__(self):
        print("------------------------------------------------------")
        print("🚀 正在初始化 OCR 模型 (CPU模式，稳定第一)...")
        print("------------------------------------------------------")
        # 强制使用 CPU，避免显卡驱动冲突
        self.reader = easyocr.Reader(['ch_sim', 'en'], gpu=False) 
        self.rows = []

    def select_images(self):
        """弹出文件选择框"""
        root = tk.Tk()
        root.withdraw()
        # 允许选择多张图片
        file_paths = filedialog.askopenfilenames(
            title="请选择【云闪付/微信/支付宝】账单截图",
            filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp")]
        )
        return file_paths

    def clean_amount(self, text):
        """从文本中提取金额 (支持 ¥1,000.00, +500 等格式)"""
        # 匹配逻辑：找 + 或 - 号，后面跟着数字
        match = re.search(r'([+\-])[¥￥]?\s*([\d,]+\.\d{2})', text)
        if match:
            sign = match.group(1)
            num = match.group(2).replace(',', '') # 去掉千分位逗号
            return sign, float(num)
        return None, None

    def parse_images(self, file_paths):
        """核心解析逻辑"""
        # 1. 初始化当前日期 (防止第一笔交易没日期)
        current_date = datetime.datetime.now().strftime("%Y-%m-%d")
        
        # 2. 常见的日期格式正则 (支持 01.15 或 1月15日)
        date_pattern = re.compile(r'(\d{1,2})[月\.](\d{1,2})') 

        for f_path in file_paths:
            print(f"\n📸 正在读取: {os.path.basename(f_path)}")
            
            try:
                # === 🚑 关键修复：支持中文路径 ===
                with open(f_path, 'rb') as f:
                    img_bytes = f.read()
                result_texts = self.reader.readtext(img_bytes, detail=0)
            except Exception as e:
                print(f"❌ 读取失败: {e}")
                continue

            # 3. 逐行分析文字
            # 这里的逻辑是：先找日期更新状态，再找金额生成记录
            for i, text in enumerate(result_texts):
                text = text.strip()
                if not text: continue

                # [A] 尝试捕捉日期 (如 "01.15")
                date_match = date_pattern.search(text)
                if date_match:
                    month, day = date_match.groups()
                    current_date = f"{DEFAULT_YEAR}-{int(month):02d}-{int(day):02d}"
                    print(f"   📅 锁定日期: {current_date}")
                    continue

                # [B] 尝试捕捉金额
                sign, amount = self.clean_amount(text)
                
                # 如果找到了金额，我们就认为这是一笔交易
                if amount is not None:
                    # [C] 智能寻找描述/名称
                    # 策略：通常描述在金额的左边或上边。OCR读出来的顺序可能是分行的。
                    # 我们先看当前行去掉金额后还有没有字
                    clean_text = re.sub(r'[+\-][¥￥]?\s*[\d,]+\.\d{2}', '', text).strip()
                    
                    description = ""
                    if len(clean_text) > 2:
                        # 当前行就有描述 (比如：收到转账 +500)
                        description = clean_text
                    elif i > 0:
                        # 当前行只有金额，描述大概率在上一行
                        description = result_texts[i-1].strip()
                    
                    # [D] 清洗描述 (去掉 "转账:", "转给" 等废话)
                    desc_final = description.replace("转账", "").replace("转给", "").replace("来自", "").replace(":", "").strip()
                    if not desc_final: desc_final = "OCR自动识别"

                    # [E] 判断收支和对象
                    record_type = "支出" if sign == '-' else "收入"
                    
                    target = ""
                    if "应翔" in desc_final: target = "应翔"
                    elif "王意帆" in desc_final: target = "王意帆"

                    # [F] 存入列表
                    self.rows.append({
                        '账户': '云闪付(待定)', # 导入 Moze 后需匹配真实账户
                        '币种': 'CNY',
                        '记录类型': record_type,
                        '主类别': '未分类', 
                        '子类别': '',
                        '金额': amount,
                        '手续费': 0, '折扣': 0,
                        '名称': desc_final,
                        '商家': '',
                        '日期': current_date,
                        '时间': '00:00:00',
                        '项目': '',
                        '描述': f"原图文字: {text}",
                        '标签': '#OCR导入',
                        '对象': target
                    })
                    print(f"      ✅ [识别成功] {current_date} | {record_type} {amount} | {desc_final}")

    def save(self):
        if not self.rows:
            print("\n⚠️ 未识别到任何有效数据，请检查图片是否清晰。")
            return

        # 获取当前脚本所在文件夹
        base_path = os.path.dirname(os.path.abspath(__file__))
        filename = f"Moze导入_{datetime.datetime.now().strftime('%H%M%S')}.csv"
        full_path = os.path.join(base_path, filename)

        try:
            df = pd.DataFrame(self.rows, columns=MOZE_COLUMNS)
            # 关键：utf-8-sig 编码，防止 Excel 和 Moze 乱码
            df.to_csv(full_path, index=False, encoding='utf-8-sig')
            
            print("\n" + "="*40)
            print(f"🎉 成功！文件已生成: {filename}")
            print(f"📂 保存在: {base_path}")
            print("="*40)
            
            # 自动打开文件夹
            os.startfile(base_path)
        except Exception as e:
            print(f"❌ 保存文件失败: {e}")

if __name__ == "__main__":
    app = MozeOCRTool()
    
    # 1. 选图
    print("等待选择图片...")
    files = app.select_images()
    
    if files:
        # 2. 识别
        app.parse_images(files)
        # 3. 保存
        app.save()
    else:
        print("用户取消了选择。")