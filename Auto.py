import os
import subprocess
import datetime
import time
from pathlib import Path

# ================= 配置区域 =================
# 这里填你的项目绝对路径 (注意前面的 r 不能删，那是为了处理 Windows 路径斜杠的)

REPO_PATH = Path(__file__).resolve().parent
# ===========================================


def run_command(cmd, description):
    """执行命令并打印美化后的输出"""
    print(f"\n[+] 正在{description}...")
    try:
        # shell=True 允许在 Windows 终端中运行命令
        # check=True 会在命令报错时抛出异常（可选）
        result = subprocess.run(cmd, shell=True, text=True)
        if result.returncode == 0:
            print(f"✅ {description}成功")
        else:
            print(f"⚠️ {description}返回了非零状态 (可能是没有变化或有冲突)")
    except Exception as e:
        print(f"❌ 执行出错: {e}")


def main():
    # 1. 检查路径是否存在
    if not os.path.exists(REPO_PATH):
        print(f"❌ 错误：找不到路径 {REPO_PATH}")
        input("按回车键退出...")
        return

    # 2. 切换到项目目录
    os.chdir(REPO_PATH)
    print(f"📂 已切换到目录: {os.getcwd()}")

    # 3. 开始 Git 三板斧

    # 第一步：拉取 (Pull)
    run_command("git pull", "从云端拉取最新代码")

    # 第二步：添加 (Add)
    run_command("git add .", "添加文件变动")

    # 第三步：提交 (Commit)
    # 生成带时间戳的备注，例如: "Auto-sync 2023-10-27 10:30:05"
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    commit_msg = f"Auto-sync {timestamp}"

    # 注意：如果暂存区没东西，git commit 会报错，这是正常的，不用管
    run_command(f'git commit -m "{commit_msg}"', "提交到本地仓库")

    # 第四步：推送 (Push)
    run_command("git push", "推送到 GitHub 和 Gitee")

    print("\n" + "="*30)
    print("🎉 同步流程结束！")
    print("="*30)

    # 暂停一下，让你看清楚结果，而不是窗口一闪而过
    input("\n按回车键关闭窗口...")


if __name__ == "__main__":
    main()
