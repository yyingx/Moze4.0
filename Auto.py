# -*- coding: utf-8 -*-
"""
Git 自动同步工具 v2.0 (优化版)
特性：
- 显示详细的 git 命令输出
- 智能检测变更，无变更时跳过 commit
- 冲突检测和错误处理
- 支持命令行参数快速模式
- commit 消息安全转义
"""

import os
import subprocess
import datetime
import sys
import re
from pathlib import Path

# ================= 配置区域 =================
REPO_PATH = Path(__file__).resolve().parent
# ===========================================


class GitAutoSync:
    def __init__(self):
        self.repo_path = REPO_PATH

    def run_command(self, cmd, description, capture=True, check_error=True):
        """执行命令并返回结果"""
        print(f"\n[+] 正在{description}...")
        try:
            if capture:
                result = subprocess.run(
                    cmd,
                    shell=True,
                    text=True,
                    capture_output=True,
                    cwd=self.repo_path
                )

                # 显示输出
                if result.stdout.strip():
                    print(result.stdout.strip())
                if result.stderr.strip():
                    print(result.stderr.strip())

                # 检查错误
                if check_error and result.returncode != 0:
                    print(f"❌ {description}失败 (返回码: {result.returncode})")
                    return False
                else:
                    print(f"✅ {description}成功")
                    return result
            else:
                result = subprocess.run(cmd, shell=True, cwd=self.repo_path)
                return result.returncode == 0

        except Exception as e:
            print(f"❌ 执行出错: {e}")
            return False

    def check_conflicts(self):
        """检查是否有合并冲突"""
        result = subprocess.run(
            "git diff --name-only --diff-filter=U",
            shell=True,
            text=True,
            capture_output=True,
            cwd=self.repo_path
        )
        conflict_files = result.stdout.strip().split(
            '\n') if result.stdout.strip() else []

        if conflict_files:
            print("\n⚠️ 发现合并冲突的文件：")
            for f in conflict_files:
                print(f"   - {f}")
            print("\n请手动解决冲突后重新运行脚本。")
            return True
        return False

    def check_changes(self):
        """检查是否有文件变更"""
        result = subprocess.run(
            "git status --porcelain",
            shell=True,
            text=True,
            capture_output=True,
            cwd=self.repo_path
        )
        return bool(result.stdout.strip())

    def show_status(self):
        """显示当前变更状态"""
        print("\n📋 当前变更文件：")
        result = subprocess.run(
            "git status --short",
            shell=True,
            text=True,
            capture_output=True,
            cwd=self.repo_path
        )
        if result.stdout.strip():
            print(result.stdout)
        else:
            print("   (无变更)")

    def escape_commit_msg(self, msg):
        """转义 commit 消息中的特殊字符"""
        # 替换双引号为单引号，避免命令注入
        msg = msg.replace('"', "'")
        return msg

    def run(self, auto_mode=False):
        """主流程"""
        # 检查路径
        if not os.path.exists(self.repo_path):
            print(f"❌ 错误：找不到路径 {self.repo_path}")
            return False

        os.chdir(self.repo_path)
        print(f"📂 工作目录: {os.getcwd()}")

        # === 1. 拉取最新代码 ===
        result = self.run_command("git pull", "从远程拉取最新代码")
        if not result:
            return False

        # 检查冲突
        if self.check_conflicts():
            return False

        # === 2. 显示当前状态 ===
        self.show_status()

        # === 3. 检查是否有变更 ===
        if not self.check_changes():
            print("\n✨ 没有新的变更需要提交")
            print("="*40)
            return True

        # === 4. 添加文件 ===
        result = self.run_command("git add .", "添加文件变动")
        if not result:
            return False

        # === 5. 提交 ===
        if auto_mode:
            # 自动模式：使用时间戳
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            commit_msg = f"Auto-sync {timestamp}"
            print(f"\nℹ️ [快速模式] 使用默认备注: {commit_msg}")
        else:
            # 交互模式：询问用户
            print("\n" + "-"*50)
            print("💡 提示：符合规范的消息格式")
            print("   新增功能: 新增 vXX.XX: 功能描述")
            print("   优化改进: 优化 vXX.XX: 改进内容")
            print("   Bug修复: 修复 vXX.XX: 问题描述")
            print("-"*50)
            user_msg = input("📝 请输入 Commit 消息 (回车=默认时间戳): ").strip()

            if not user_msg:
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                commit_msg = f"Auto-sync {timestamp}"
                print(f"ℹ️ 使用默认备注: {commit_msg}")
            else:
                commit_msg = user_msg

        # 转义消息
        commit_msg = self.escape_commit_msg(commit_msg)

        # 执行提交
        result = self.run_command(
            f'git commit -m "{commit_msg}"',
            "提交到本地仓库",
            check_error=True
        )
        if not result:
            return False

        # === 6. 推送 ===
        result = self.run_command("git push", "推送到远程仓库")
        if not result:
            return False

        # 成功提示
        print("\n" + "="*40)
        print("🎉 同步流程成功完成！")
        print("="*40)
        return True


def main():
    """主入口"""
    # 检查命令行参数
    auto_mode = len(sys.argv) > 1 and sys.argv[1] in [
        '-a', '--auto', '-f', '--fast']

    if auto_mode:
        print("🚀 [快速模式] 自动使用时间戳作为 commit 消息\n")

    syncer = GitAutoSync()
    success = syncer.run(auto_mode=auto_mode)

    # 如果不是快速模式，等待用户确认
    if not auto_mode:
        input("\n按回车键关闭窗口...")

    # 返回退出码
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
