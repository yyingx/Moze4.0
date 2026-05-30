---
name: commit
description: 按项目规范提交代码
disable-model-invocation: true
allowed-tools: Bash, Read
---

# Moze 项目 Commit Skill

根据项目 commit 规范提交代码变更。

## Commit 格式

```
<类型> <版本/模块>: <描述>

Co-Authored-By: Codex <model> <noreply@anthropic.com>
```

**类型：**
- `新增` - 新功能
- `优化` - 改进现有功能
- `修复` - Bug 修复

**版本/模块示例：**
- `v11.84` - Import 脚本版本
- `云闪付PaddleOCR_lite稳定版` - 模块名称

## 执行步骤

1. 运行 `git status` 查看变更文件（不使用 -uall）
2. 运行 `git diff --staged` 和 `git diff` 查看具体改动
3. 运行 `git log --oneline -5` 参考最近 commit 风格
4. 分析变更，草拟符合规范的 commit message
5. 向用户确认 commit message
6. 用户确认后执行 `git add` 和 `git commit`

## 注意事项

- 不要自动 push，除非用户明确要求
- 不要提交 .env、credentials 等敏感文件
- 不要提交支付宝/微信源账单、Moze 导出 CSV、截图账单等私人数据
- 使用 HEREDOC 格式传递 commit message 以保证格式正确
- 不强制添加 Co-Author；如项目需要，再按当前工具身份添加
