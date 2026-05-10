import os
import tkinter as tk
from tkinter import filedialog
from pydub import AudioSegment
from pydub.utils import mediainfo  # 新增：导入媒体信息分析工具


def get_input_file():
    # 初始化一个隐藏的图形窗口
    root = tk.Tk()
    root.withdraw()

    print("等待选择文件...")

    # 弹出文件选择对话框
    file_path = filedialog.askopenfilename(
        title="请选择你从 YouTube 下载的音频文件",
        filetypes=[("音频文件", "*.mp3 *.m4a *.webm *.wav *.aac"), ("所有文件", "*.*")]
    )
    return file_path


def convert_for_car():
    # 1. 弹出窗口让用户选择文件
    input_file = get_input_file()

    # 如果用户点了取消或关掉了窗口，直接退出
    if not input_file:
        print("❌ 你取消了文件选择，程序已退出。")
        return

    print(f"\n📁 已选择文件: {input_file}")

    # 2. 自动生成输出文件名
    file_dir, original_name = os.path.split(input_file)
    name_only, _ = os.path.splitext(original_name)
    output_file = os.path.join(file_dir, f"{name_only}_FordReady.mp3")

    # ================= 新增：读取并打印原音频信息 =================
    print("\n🔍 正在分析源文件真实编码信息...")
    try:
        info = mediainfo(input_file)

        # 提取格式、采样率、比特率和时长
        format_name = info.get('format_name', '未知')
        sample_rate = info.get('sample_rate', '未知')

        bit_rate = info.get('bit_rate')
        if bit_rate and bit_rate.isdigit():
            bit_rate_kbps = f"{int(bit_rate) // 1000} kbps"
        else:
            bit_rate_kbps = "未知/VBR动态码率"

        duration = info.get('duration')
        if duration:
            mins = int(float(duration)) // 60
            secs = int(float(duration)) % 60
            duration_str = f"{mins}分{secs}秒"
        else:
            duration_str = "未知"

        print("-" * 40)
        print(f"📄 真实底层格式: {format_name}")
        print(f"⏱️ 音频总时长:   {duration_str}")
        print(f"🎛️ 原始采样率:   {sample_rate} Hz (车机通常只认 44100 Hz)")
        print(f"📊 原始比特率:   {bit_rate_kbps} (车机通常推荐 128~192 kbps CBR)")
        print("-" * 40)

    except Exception as e:
        print(f"⚠️ 无法读取文件信息，可能是文件损坏或 ffmpeg 未配置好: {e}")
    # ==========================================================

    print("\n🎧 正在读取并准备重编码...")

    try:
        # 3. 读取音频
        audio = AudioSegment.from_file(input_file)

        print("⚙️ 读取成功！正在重编码为车机标准 MP3 (44100Hz, 192kbps CBR) ...")

        # 4. 强制导出为车机兼容格式
        audio.export(output_file,
                     format="mp3",
                     bitrate="192k",
                     parameters=["-ar", "44100"])

        print("\n=========================================")
        print(f"✅ 转换大功告成！")
        print(f"🎵 你的新文件存放在: {output_file}")
        print("🚗 赶紧把它拷进 U 盘插到福特车上试试吧！")
        print("=========================================")

    except FileNotFoundError as e:
        print(f"\n❌ 运行错误: {e}")
        print("💡 提示: pydub 底层依赖 ffmpeg。如果看到此报错，请检查电脑是否已安装 ffmpeg 并配置了系统环境变量。")
    except Exception as e:
        print(f"\n❌ 发生未知错误: {e}")


if __name__ == "__main__":
    convert_for_car()
