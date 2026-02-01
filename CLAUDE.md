# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 概述

Python 脚本集，将支付宝/微信/云闪付账单转换为 Moze 4.0 导入格式。

## 快速开始

```bash
# 依赖安装
pip install pandas numpy openpyxl
pip install paddlepaddle==2.6.2 paddleocr==2.7.3 opencv-python  # OCR 需要

# 常用命令
python Moze4.0_Import_v11_73.py           # 支付宝/微信账单转换（推荐）
python "Moze4.0_云闪付PaddleOCR 稳定版_.py"  # 云闪付 OCR
python Moze4.0报销_Enhanced.py             # 差旅报销生成
python Auto.py                             # Git 同步（交互模式）
python Auto.py -a                          # Git 同步（快速模式）
```

## 版本选择

| 版本 | 特性 | 推荐场景 |
|------|------|----------|
| **v11.73** | 扩展手动添加行关键词（56个子类别+15个名称） | **日常使用** |
| v11.71 | 动态 header 检测、异常处理、logging | 稳定版 |
| v11.70 | JSON 配置分离 (`ingredient_config.json`) | 配置管理 |

## 核心数据结构

**⚠️ 修改前必读：**

- **`DATA_SOURCE`** - 关键词字典（不要随便改）
- **`INGREDIENT_PRIORITY`** - 匹配优先级顺序（顺序决定匹配结果）
- **`RAW_MAPPING_CONFIG`** - 子类别 → (记录类型, 主类别, 项目) 映射
- **`SUBCAT_KEYWORDS`** - 备注中的子类别关键词（v11.73 扩展到 56 个）
- **`NAME_KEYWORDS`** - 备注中的名称关键词（v11.73 扩展到 15 个）

**修改禁忌：**
- 不要打乱 `INGREDIENT_PRIORITY` 顺序
- 不要删除 `RAW_MAPPING_CONFIG` 中已有的子类别
- 不要在 `exclude_if_match` 中引用不存在的 key

## 手动添加行语法

在 Excel 备注列使用 `关键词+描述` 格式：

| 备注 | 解析结果 |
|------|----------|
| `日用洗衣液` | 子类别=日常用品, 描述=洗衣液 |
| `药品感冒药` | 子类别=药品, 描述=感冒药 |
| `图书Python` | 子类别=图书, 描述=Python |
| `正餐麦当劳` | 根据时间推导早/午/晚餐, 描述=麦当劳 |
| `借入谢辉.钉子` | 子类别=借入, 对象=谢辉, 描述=钉子 |
| `住宿费汉庭` | 子类别=报账, 名称=住宿费, 对象=天之逸, 标签=#差旅报销 |

**v11.73 支持的子类别关键词（56个）：**
- 饮食：日用、食材、零食、饮料水果、纯净水、早餐、午餐、晚餐、夜宵
- 购物：服饰、数码、家具、大件
- 交通：加油、充电、公共交通、火车、机票、出租车、停车费
- 居家：房租、水费、电费、物业费、宽带费、快递、理发、电话费
- 虚拟：App、Software、软件、订阅、影音
- 娱乐：电影、聚会、旅游、游戏
- 医疗：药品、门诊、体检、保健用品
- 学习：图书、教材、探索、文具
- 个人：保险、礼金、红包
- 收入：薪资、福利补贴、年终奖

## 关键文件

- **`Moze Dict.xlsx`** - 商家规则字典（必需）
- **`ingredient_config.json`** - 分类配置（仅 v11.70）
- **`Moze4.0_Import/`** - 输出目录

## 版本管理

**命名规则：** `Moze4.0_Import_v11_XX.py`

**Commit 规范：**
```
新增 vXX.XX: 功能描述
优化 vXX.XX: 改进内容
修复 vXX.XX: 问题描述
```

## 报销生成器配置

```python
CFG = {
    'MEAL': 40,      # 餐补（元/天）
    'TRANS': 20,     # 交补（元/天）
    'HOTEL': 140,    # 住宿标准（元/天）
}
```
日期格式：`yymmdd` 或 `yymmdd-yymmdd`

## 云闪付 OCR 配置

```python
DEFAULT_YEAR = 2026
MY_NAME = "应翔"
ACCOUNT_MAP = {...}  # 银行卡尾号映射
```

## 调试

```python
# 启用详细日志
logging.basicConfig(level=logging.DEBUG)

# 测试单个文件
file_path = Path("测试文件.csv")
```
