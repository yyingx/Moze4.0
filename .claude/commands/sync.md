# Git 同步

运行 Git 自动同步脚本，提交并推送代码变更。

## 执行命令

```bash
python Auto.py $ARGUMENTS
```

## 说明

- 无参数：交互模式，可编辑 commit message
- `-a` 参数：快速模式，自动生成 commit message 并推送

## 参数示例

- `/sync` - 交互模式
- `/sync -a` - 快速自动同步
