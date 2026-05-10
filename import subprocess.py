import subprocess

# 你的源文件和目标文件
input_file = "你下载的音频.webm"
output_file = "Ford_Car_Ready.mp3"

# 这就是 Python 要在后台敲打的命令，拆解成了列表形式
command = [
    "ffmpeg",                 # 呼叫 ffmpeg 程序
    "-y",                     # 如果输出文件已存在，直接覆盖不提示
    "-i", input_file,         # 指定输入文件
    "-b:a", "192k",           # 强制音频码率为 192k CBR
    "-ar", "44100",           # 强制采样率为 44100Hz
    output_file               # 指定输出文件名
]

print("Python 正在后台使唤 ffmpeg 干活...")

# subprocess.run 就是那个“执行键”，它会调起系统终端来跑上面的命令
# 跑完之前，Python 脚本会停在这里等它
result = subprocess.run(command)

if result.returncode == 0:
    print("干活完毕！")
else:
    print("ffmpeg 报错了！")
