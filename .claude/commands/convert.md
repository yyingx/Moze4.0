# 账单转换

运行 Moze 4.0 账单转换脚本，将支付宝/微信账单转换为 Moze 导入格式。

## 执行命令

```bash
python "$(ls -1 Moze4.0_Import_v*.py | grep -Ev 'experimental|预发行' | sort -V | tail -1)" $ARGUMENTS
```

## 说明

- 脚本会自动检测账单类型（支付宝 CSV / 微信 XLSX）
- 输出文件保存在 `Moze4.0_Import/` 目录
- 如需指定文件，在 `/convert` 后加文件路径

## 参数示例

- `/convert` - 交互式选择文件
- `/convert 支付宝账单.csv` - 指定文件转换
