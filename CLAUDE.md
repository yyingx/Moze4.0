# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 概述

本仓库包含 Python 脚本，用于将支付宝 CSV 和微信 XLSX 账单转换为 Moze 4.0 导入格式，实现自动记账。

## 运行脚本

```bash
python Moze4_0_Import_v11_66_8.py
```

脚本会弹出文件选择对话框，选择账单文件后输出 CSV 到 `Moze4.0_Import/` 目录。

**依赖安装：**
```bash
pip install pandas numpy openpyxl
```

## 架构

### 主脚本 (`Moze4_0_Import_v11_66_8.py`)

处理流程：
1. `load_settings()` / `load_rules()` - 从 `Moze Dict.xlsx` 加载配置和规则
2. `sniff_and_load_data()` - 检测文件类型（CSV=支付宝，XLSX=微信）并解析
3. `process_transfers()` - 处理账户间转账交易
4. `process_main()` - 主交易处理，匹配字典规则
5. `process_heuristics()` - 基于 `DATA_SOURCE` 关键词的自动分类
6. `save_result()` - 输出带时间戳的 CSV

### 核心数据结构

**请勿随意修改以下结构，它们相互依赖：**

- `DATA_SOURCE` - 各类别的关键词列表（水果、蔬菜、肉类等）
- `INGREDIENT_PRIORITY` - 食材分类处理顺序：`(key, 名称, 子类别)`
- `RAW_MAPPING_CONFIG` - 子类别映射到 `(记录类型, 主类别, 项目)`
- `AUTO_MAP_DICT` - 由 `RAW_MAPPING_CONFIG` 派生，子类别 → 元组

`DATA_SOURCE` 和 `INGREDIENT_PRIORITY` 的顺序必须对应。

### 分类优先级

1. **备注驱动** - 用户在备注中标注，如 `探索claude` → 子类=探索，描述=claude
2. **字典规则** (`Moze Dict.xlsx`) - 先精确匹配商家名，再正则匹配
3. **关键词检测** (`DATA_SOURCE`) - 在商家/商品名中检测食材/商品关键词
4. **时间推导餐饮** - 检测为正餐但无子类别时，根据交易时间判断早餐/午餐/晚餐/夜宵

### 淘宝订单识别

通过商户单号前缀 `T200P` 识别淘宝订单（兼容 T200P4、T200P5 等）。

### 转账处理

通过 `CONFIG['TRANSFER_TARGET_x']` 配置转账对方，自动生成转入/转出配对记录。

## 关键文件

- `Moze4_0_Import_v11_66_8.py` - 当前稳定版本
- `Moze Dict.xlsx` - 规则文件，商家→分类映射（必需）
- `Auto.py` - Git 自动同步工具（pull, add, commit, push）

## 备注语法

在支付宝/微信备注中使用格式：`关键词+描述`

| 备注 | 结果 |
|------|------|
| `探索claude` | 子类=探索，描述=claude |
| `日用洗衣液` | 子类=日常用品，描述=洗衣液 |
| `正餐麦当劳` | 时间推导→早餐/午餐/晚餐，描述=麦当劳 |
| `借出张三` | 子类=借出，对象=张三 |
