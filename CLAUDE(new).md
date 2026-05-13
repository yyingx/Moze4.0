# CLAUDE.md — Moze 导入脚本项目指引

> 适用版本：v11.76+  
> 作者：TZY_YX  
> 本文件供 Claude Code 快速理解项目结构、数据流、修改约定及已知陷阱。

---

## 一、项目概述

将**支付宝（CSV）/ 微信支付（XLSX）** 账单转换为 **Moze 4.0 可导入的 CSV 格式**。

核心任务：
- 清洗原始账单字段
- 通过字典规则 + 启发式关键词，推导出 `记录类型 / 主类别 / 子类别 / 商家 / 描述` 等字段
- 输出到 `Moze4.0_Import/MOZE导入_时间戳.csv`

---

## 二、文件结构

```
项目根目录/
├── Moze4_0_Import_v11.76.py   # 主脚本（唯一入口）
├── Moze Dict.xlsx              # 规则字典（必须存在）
│   ├── sheet: Moze Dict        # 匹配规则表（商家/商品 → 分类/描述等）
│   └── sheet: Settings         # 可覆盖 CONFIG 的自定义配置
└── Moze4.0_Import/             # 输出目录（自动创建）
    └── MOZE导入_YYYYMMDD_HHMMSS.csv
```

---

## 三、处理管道（数据流）

```
账单文件（.csv / .xlsx）
    │
    ▼
sniff_and_load_data()          # 动态检测 header 行，统一列名，标记来源
    │
    ▼
robust_date_converter()        # 转换日期格式
    │
    ├─── 转账对方？─── YES ──▶ process_transfers()   # 生成转入/转出双条记录
    │
    └─── NO
          │
          ▼
       process_main()
          ├─ 过滤无效状态 & 零金额
          ├─ 自动推断空白的 收/支 列
          ├─ 过滤：只保留支出 | 收入有关键词 | 备注有特殊词
          ├─ ensure_columns()           # 补齐所有必要列
          ├─ apply_rules()              # 字典匹配（简单规则 + 复合规则）
          ├─ construct_description()    # 构建描述（备注 > 字典描述）
          ├─ parse_memo_subcategory()   # 解析备注中的子类别前缀
          └─ process_heuristics()       # 启发式分类推导
                ├─ _apply_ingredient_patterns()   # 食材/商品词匹配
                ├─ process_debt_keywords()        # 借贷关键词
                ├─ process_reimbursement()        # 报销关键词
                ├─ process_location_tags()        # 地点标签
                ├─ process_generic_keywords()     # 通用分类词
                ├─ _infer_meal_time()             # 正餐时间段推导
                └─ _map_record_types()            # 子类 → 记录类型/主类/项目
    │
    ▼
finalize_records()             # 金额符号、日期时间、账户标准化
    │
    ▼
save_result()                  # 排序、打标签、写 CSV
```

---

## 四、关键数据结构

### 4.1 Moze Dict.xlsx — 规则字典列结构

| 列名 | 说明 |
|------|------|
| `商家(old)` | 匹配来源：交易对方（精确或正则） |
| `商品` | 可选。有值时为**复合规则**，同时匹配商家+商品 |
| `is_regex` | 0=精确匹配，1=正则匹配 |
| `商家` | 推导出的商家名 |
| `主类别` | 主类别（饮食/购物/交通…） |
| `子类别` | 子类别（午餐/日常用品…） |
| `描述` | 描述内容（可选） |
| *(其他列)* | 其余列均写入输出 CSV |

> `商家(old)` / `is_regex` / `商品` 是过滤列，不进入输出。

### 4.2 CONFIG 可覆盖项

`Settings` sheet 中 key-value 对可覆盖 `CONFIG` 字典里的任意值：
- 转账对方名称（`TRANSFER_TARGET_1/2/SNOWBALL`）
- 账户名称（`ACCOUNT_*`）

### 4.3 AUTO_MAP_DICT

由 `RAW_MAPPING_CONFIG` 在模块加载时自动展开，结构为：

```python
# 子类别 → (记录类型, 主类别, 项目)
AUTO_MAP_DICT['午餐'] = ('支出', '饮食', '食')
AUTO_MAP_DICT['薪资'] = ('收入', '收入', '工作')
```

---

## 五、描述字段（描述）优先级

**规则：备注 > 字典描述，字典描述只在备注为空时填充。**

```
construct_description() 执行顺序：
  1. 描述 = ""（清空）
  2. 淘宝订单 → 描述 = 商品名
  3. 备注非空 → 描述 = 备注内容（覆盖步骤2）
  4. 描述仍为空 → 描述 = 描述_rule（字典描述）
```

**重要**：`描述_rule` 是字典规则写入描述时使用的临时列名，不是直接写 `描述`。

---

## 六、规则匹配机制（apply_rules）

### 简单规则（只匹配 商家(old)）
```
精确匹配 → pd.merge（批量，高效）
正则匹配 → 逐行 str.contains（仅对未命中行）
结果写入 xxx_rule 临时列 → combine_first 合并回原列
描述_rule 单独保留，交给 construct_description 处理
```

### 复合规则（同时匹配 商家(old) + 商品）
```
逐行 str.contains 双条件匹配
优先级高于简单规则
描述 → 写入 描述_rule（不直接写 描述，否则被 construct_description 清空）
商家 → 仅在简单规则未命中时写入（避免覆盖已推导的干净商家名）
其他列 → 直接写入原列
```

> ⚠️ **陷阱**：复合规则的描述如果直接写 `df['描述']`，会被 `construct_description` 首行的 `df['描述'] = ""` 清空。必须写入 `描述_rule`。

---

## 七、收入记录的过滤逻辑（v11.76）

收入记录不能无条件保留，否则大量未分类收入混入。过滤规则：

```python
保留条件 = (
    收/支 == '支出'                          # 所有支出
    | (收/支 == '收入' & 备注命中收入关键词)   # 收入需有关键词
    | 备注命中特殊关键词                       # 借贷/报销等特殊词
)
```

收入关键词（`INFER_INCOME_PATTERN`）：
`薪资 / 福利补贴 / 年终奖 / 收红包 / 利息收入 / 投资盈利 / 二手折旧 / 其他收入`

---

## 八、账户标准化

支付方式字符串通过 `STANDARDIZE_ACCOUNTS` 正则批量替换：

```python
STANDARDIZE_ACCOUNTS = {
    r'.*4946.*': '平安银行4946',
    r'.*9579.*': '工商银行',
    r'.*3379.*': '招商银行Ⅱ',
    r'.*4826.*': '广发银行4826',
    r'.*零钱.*': '零钱3'
}
```

---

## 九、新增/修改规则指南

### 新增商家规则
在 `Moze Dict.xlsx` → `Moze Dict` sheet 中增加一行：
- 简单规则：填写 `商家(old)` + 目标字段，`商品` 留空
- 复合规则：同时填写 `商家(old)` + `商品`

### 新增子类别
1. 在 `RAW_MAPPING_CONFIG` 中将子类别加入对应的 `(记录类型, 主类别, 项目)` 分组
2. 如需备注推导，在 `SUBCAT_KEYWORDS` 中添加关键词映射
3. 如需商品名推导，在 `NAME_KEYWORDS` 中添加
4. 如需食材类商品词匹配，在 `DATA_SOURCE` + `INGREDIENT_PRIORITY` 中添加

### 新增食材词
```python
# DATA_SOURCE 中加入词列表
DATA_SOURCE['NEW_KEY'] = ["词A", "词B"]

# INGREDIENT_PRIORITY 中加入匹配规则
# 格式: (key, 名称, 子类别) 或 (key, 名称, 子类别, [排除key列表])
INGREDIENT_PRIORITY.append(('NEW_KEY', '名称', '食材'))
```

---

## 十、已知陷阱 & 历史 Bug

| 版本 | 问题 | 修复位置 |
|------|------|---------|
| v11.76 | 收入记录被全部过滤 | `process_main` 过滤行加 `mask_income_with_keyword` |
| v11.76 | 收入类备注无法推断收/支 | 新增 `INFER_INCOME_PATTERN` |
| v11.76 | 复合规则描述字段为空 | 复合规则写 `描述_rule` 而非 `描述` |
| v11.76 | 复合规则商家无优先级控制 | 商家仅在简单规则未命中时由复合规则写入 |
| v11.74 | `process_main` 重复定义 `all_reim_keywords` | 提升为模块级 `ALL_REIM_KEYS` |

---

## 十一、模块级预编译常量（不要在函数内重建）

以下正则/字典在模块加载时一次性构建，**函数内直接使用，禁止重建**：

| 常量 | 用途 |
|------|------|
| `PATTERNS` | DATA_SOURCE 各分类的正则（dict） |
| `AUTO_MAP_DICT` | 子类别 → (记录类型, 主类, 项目) |
| `DEBT_PATTERN` | 借贷关键词提取 |
| `REIM_PATTERN` | 报销关键词提取 |
| `GENERIC_PATTERN` | 通用分类词匹配 |
| `ALL_SPECIAL_PATTERN` | 备注特殊词过滤（不含收入词） |
| `INFER_SUBCAT_PATTERN` | 推断支出收/支 |
| `INFER_NAME_PATTERN` | 推断支出收/支（NAME_KEYWORDS） |
| `INFER_INCOME_PATTERN` | 推断收入收/支 |

---

## 十二、输出字段顺序

由 `Moze Dict.xlsx` 的列顺序决定（排除 `商家(old)` / `is_regex` / `商品`）。  
脚本中 `cols` 变量在 `main()` 里从规则文件动态读取，不硬编码。
