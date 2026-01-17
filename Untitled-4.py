# -*- coding: utf-8 -*-
"""
Moze 4.0 OCR 侦探版 (专门用来抓 Bug)
功能：强制打印 EasyOCR 识别到的每一行原始文字
"""
import os

# 🚑 防报错补丁
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import easyocr
import tkinter as tk
from tkinter import filedialog

class BColors:
    HEADER = '\033[95m'; OKBLUE = '\033[94m'; OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'; WARNING = '\033[93m'; FAIL = '\033[91m'
    ENDC = '\033[0m'; BOLD = '\033[1m'; UNDERLINE = '\033[4m'

def debug_run():
    print(f"{BColors.HEADER}{'='*60}")
    print(f"🕵️‍♂️ OCR 侦探模式启动... (准备抓出那个读错的字)")
    print(f"{'='*60}{BColors.ENDC}")

    # 1. 初始化
    print("正在加载模型 (GPU: False)...")
    reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)

    # 2. 选图
    root = tk.Tk(); root.withdraw()
    file_paths = filedialog.askopenfilenames(title="请选择那张识别失败的图片")
    
    if not file_paths: print("❌ 取消选择"); return

    for f_path in file_paths:
        print(f"\n{BColors.OKCYAN}📸 正在深度分析: {os.path.basename(f_path)}{BColors.ENDC}")
        print("-" * 50)
        
        try:
            with open(f_path, 'rb') as f: img_bytes = f.read()
            # detail=0 只返回文字
            texts = reader.readtext(img_bytes, detail=0)
            
            # 3. 打印每一行原始数据 (这就是 AI 眼里的世界)
            print(f"{BColors.WARNING}👇👇👇 请把下面这些内容发给我 👇👇👇{BColors.ENDC}\n")
            
            for i, line in enumerate(texts):
                # 打印行号和内容，方便我们定位
                print(f"[第{i:02d}行] {line}")
                
            print(f"\n{BColors.WARNING}👆👆👆 复制上面这些内容 👆👆👆{BColors.ENDC}")
            print("-" * 50)
            
        except Exception as e:
            print(f"❌ 发生意外错误: {e}")

if __name__ == "__main__":
    debug_run()
    input(f"\n{BColors.OKBLUE}按回车键退出...{BColors.ENDC}")