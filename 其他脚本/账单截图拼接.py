import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageOps
import os
import datetime
import math

def create_smart_layout_pdf():
    # 1. 交互式选择图片
    root = tk.Tk()
    root.withdraw()
    print("请选择要拼接的图片（支持多选）...")
    file_paths = filedialog.askopenfilenames(
        title="选择截图文件",
        filetypes=[("Image files", "*.jpg *.jpeg *.png *.bmp *.webp")]
    )

    if not file_paths:
        print("未选择图片，程序结束。")
        return

    images = []
    base_dir = os.path.dirname(file_paths[0]) if file_paths else ""

    print("正在加载图片...")
    for fp in file_paths:
        try:
            img = Image.open(fp)
            img = ImageOps.exif_transpose(img)
            images.append(img)
        except Exception as e:
            print(f"无法加载图片 {fp}: {e}")

    count = len(images)
    if count == 0:
        return

    # 2. 定义横向 A4 尺寸 (300 DPI)
    A4_WIDTH = 3508
    A4_HEIGHT = 2480
    
    canvas = Image.new('RGB', (A4_WIDTH, A4_HEIGHT), (255, 255, 255))

    # 3. 智能排版逻辑
    # 如果图片 >= 6 张，强制分 2 行
    # 如果图片 < 6 张，保持 1 行
    if count >= 6:
        rows = 2
        # 列数向上取整。比如 6张->3列，7张->4列
        cols = math.ceil(count / rows)
        print(f"图片较多 ({count}张)，自动采用 {rows}行 x {cols}列 排版...")
    else:
        rows = 1
        cols = count
        print(f"图片较少 ({count}张)，采用单行排版...")
    
    # 计算每个格子的大小
    cell_width = A4_WIDTH // cols
    cell_height = A4_HEIGHT // rows

    for i, img in enumerate(images):
        # 计算当前图片处于第几行、第几列
        row_idx = i // cols
        col_idx = i % cols
        
        x_start = col_idx * cell_width
        y_start = row_idx * cell_height

        # 4. 调整图片大小以适应格子
        img_w, img_h = img.size
        # 核心缩放算法：取宽比和高比中较小的，保证完全放入
        ratio = min(cell_width / img_w, cell_height / img_h)
        
        new_w = int(img_w * ratio)
        new_h = int(img_h * ratio)
        
        # 高质量缩放
        resized_img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
        
        # 5. 居中粘贴
        paste_x = x_start + (cell_width - new_w) // 2
        paste_y = y_start + (cell_height - new_h) // 2
        
        canvas.paste(resized_img, (paste_x, paste_y))

    # 6. 保存 PDF
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"A4拼图_{count}张_{timestamp}.pdf"
    save_path = os.path.join(base_dir, output_filename)

    print(f"正在保存 PDF 至: {save_path} ...")
    
    try:
        canvas.save(save_path, "PDF", resolution=300.0)
        print("✅ 成功！文件已保存。")
        try:
            os.startfile(save_path)
        except AttributeError:
            pass
    except Exception as e:
        print(f"❌ 保存失败: {e}")

if __name__ == "__main__":
    create_smart_layout_pdf()