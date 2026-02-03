# 云闪付 OCR

运行云闪付账单 OCR 识别脚本，从截图提取交易记录并转换为 Moze 导入格式。

## 执行命令

```bash
python "Moze4.0_云闪付PaddleOCR 稳定版_.py" $ARGUMENTS
```

## 说明

- 使用 PaddleOCR 识别云闪付截图
- 自动匹配银行卡尾号到账户
- 输出文件保存在 `Moze4.0_Import/` 目录

## 参数示例

- `/ocr` - 交互式选择图片
- `/ocr 截图.jpg` - 指定图片识别
