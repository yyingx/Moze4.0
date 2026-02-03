# 报销单生成

运行差旅报销生成器，根据日期范围生成报销记录。

## 执行命令

```bash
python Moze4.0报销_Enhanced.py $ARGUMENTS
```

## 说明

- 日期格式：`yymmdd` 或 `yymmdd-yymmdd`（日期范围）
- 默认补贴标准：餐补 40 元/天、交补 20 元/天、住宿 140 元/天
- 输出文件保存在 `Moze4.0_Import/` 目录

## 参数示例

- `/reimburse 260115` - 单日报销
- `/reimburse 260115-260118` - 日期范围报销
