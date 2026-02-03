---
name: "moze-rule-expert"
description: "Moze 4.0 账单匹配规则与 ETL 逻辑专家。在修改映射字典、审计分类或更新交易处理规则时调用。"
---

# Moze 规则专家 (Moze Rule Expert)

此技能为维护 Moze 4.0 账单导入脚本及相关映射数据提供专业支持。

## 核心能力
- **关键词冲突检测**：扫描 `DATA_SOURCE` 和 `INGREDIENT_PRIORITY`，确保没有重叠的关键词导致分类错误。
- **映射验证**：验证关键词中定义的所有子类别在 `RAW_MAPPING_CONFIG` 中都有有效的映射（记录类型、主类别、项目）。
- **商家规则审计**：通过建议正则改进和防止重复规则，帮助维护 `Moze Dict.xlsx`。
- **合规性检查**：确保所有改动符合 [CLAUDE.md](file:///e:/天之逸2025/Moze4.0/CLAUDE.md) 中列出的“修改禁忌”。

## 何时调用
- **规则更新**：在 [Moze4.0_Import_v11_74.py](file:///e:/天之逸2025/Moze4.0/Moze4.0_Import_v11_74.py) 中添加或修改关键词时。
- **分类变更**：当 [CATEGORY_STRUCTURE.md](file:///e:/天之逸2025/Moze4.0/CATEGORY_STRUCTURE.md) 中的 Moze 4.0 分类树更新时。
- **故障排除**：当某笔交易被错误分类或过滤时。

## 执行流程 (Workflow)
当你需要修改或审计规则时，**必须**按以下步骤操作：

### 第一步：冲突预检
- 搜索 `DATA_SOURCE`：确认新关键词是否已存在于其他分类中。
- **动作**：使用 `grep` 或 `SearchCodebase` 全文检索该关键词。

### 第二步：优先级评估
- 检查 `INGREDIENT_PRIORITY`：新分类应该放在什么位置？
- **原则**：更具体的词（如“土鸡蛋”）必须先于通用词（如“鸡蛋”）匹配。

### 第三步：映射链审计
- 检查 `RAW_MAPPING_CONFIG`：
    1. 确保主类别存在于 [CATEGORY_STRUCTURE.md](file:///e:/天之逸2025/Moze4.0/CATEGORY_STRUCTURE.md)。
    2. 确保记录类型（支出/收入/转账）正确。
- **动作**：对比脚本中的 `RAW_MAPPING_CONFIG` 字典。

### 第四步：代码同步检查
- 检查 [Moze4.0_Import_v11_74.py](file:///e:/天之逸2025/Moze4.0/Moze4.0_Import_v11_74.py)：
    1. 如果增加了 `SUBCAT_KEYWORDS`，必须同步更新 `NAME_KEYWORDS`（如有必要）。
    2. 运行 `py_compile` 验证语法。

## 关键约束
- **不要** 打乱 `INGREDIENT_PRIORITY` 的顺序。
- **不要** 在没有迁移计划的情况下从 `RAW_MAPPING_CONFIG` 中删除现有的子类别。
- **务必** 验证新关键词不会与现有的高优先级分类冲突。

## 参考文件
- [CLAUDE.md](file:///e:/天之逸2025/Moze4.0/CLAUDE.md)：通用指南和修改禁忌。
- [Moze4.0_Import_v11_74.py](file:///e:/天之逸2025/Moze4.0/Moze4.0_Import_v11_74.py)：当前活跃的导入脚本。
- `Moze Dict.xlsx`：特定商家的映射规则。
