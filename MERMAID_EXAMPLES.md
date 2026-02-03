# Mermaid 图表示例

本文档展示 Moze 4.0 账单转换工具的核心流程和数据结构。

## 1. 整体工作流程

```mermaid
flowchart TD
    A[原始账单] --> B{账单类型}
    B -->|支付宝/微信| C[v11.73 脚本]
    B -->|云闪付| D[OCR 识别]
    B -->|Excel 手动| E[手动添加行]

    C --> F[读取 Moze Dict.xlsx]
    D --> F
    E --> F

    F --> G[关键词匹配]
    G --> H[DATA_SOURCE 查询]
    G --> I[备注关键词解析]

    H --> J[生成 Moze 记录]
    I --> J

    J --> K[输出 CSV]
    K --> L[Moze 4.0 导入]

    style A fill:#e1f5ff
    style L fill:#c8e6c9
    style F fill:#fff9c4
```

## 2. 关键词匹配优先级

```mermaid
graph LR
    A[交易备注] --> B{优先级匹配}

    B -->|1| C[SUBCAT_KEYWORDS<br/>60个子类别]
    B -->|2| D[NAME_KEYWORDS<br/>15个名称]
    B -->|3| E[应收应付关键词<br/>借出/借入/代付]
    B -->|4| F[报销关键词<br/>差旅/费用]

    C --> G[确定子类别]
    D --> H[确定子类别+名称]
    E --> I[确定对象]
    F --> J[确定对象+标签]

    G --> K[RAW_MAPPING_CONFIG]
    H --> K
    I --> K
    J --> K

    K --> L[最终记录类型+主类别+项目]

    style B fill:#ffccbc
    style K fill:#c5cae9
```

## 3. 数据结构关系

```mermaid
classDiagram
    class DataSource {
        +Dict~str,List~ keywords
        +match(transaction) str
    }

    class IngredientConfig {
        +str subcategory
        +str record_type
        +str main_category
        +str project
        +List exclude_conditions
    }

    class Transaction {
        +str remark
        +float amount
        +datetime time
        +str merchant
    }

    class MozeRecord {
        +str record_type
        +str main_category
        +str subcategory
        +str project
        +str name
        +str description
        +float amount
        +str account
    }

    DataSource --> Transaction : 匹配
    Transaction --> IngredientConfig : 查询映射
    IngredientConfig --> MozeRecord : 生成

    class ManualKeywords {
        +Dict SUBCAT_KEYWORDS
        +Dict NAME_KEYWORDS
        +List EXPENSE_KEYWORDS
    }

    ManualKeywords --> IngredientConfig : 覆盖
```

## 4. 手动添加行处理流程

```mermaid
sequenceDiagram
    participant U as 用户
    participant E as Excel备注列
    participant P as 解析器
    participant M as 映射器
    participant O as 输出

    U->>E: 输入 "日用洗衣液"
    E->>P: 传递备注文本

    P->>P: 检测 "日用" 关键词
    P->>M: 查询 SUBCAT_KEYWORDS
    M-->>P: 返回子类别="日常用品"

    P->>P: 提取描述="洗衣液"
    P->>M: 查询 RAW_MAPPING_CONFIG
    M-->>P: 返回 (支出,购物,-)

    P->>O: 生成记录
    Note over O: 子类别=日常用品<br/>描述=洗衣液<br/>主类别=购物

    O->>U: 写入 Moze CSV
```

## 5. 时间推导逻辑

```mermaid
stateDiagram-v2
    [*] --> 检测正餐关键词

    检测正餐关键词 --> 获取交易时间

    获取交易时间 --> 早餐: 05:00-10:00
    获取交易时间 --> 午餐: 10:00-15:00
    获取交易时间 --> 晚餐: 15:00-21:00
    获取交易时间 --> 夜宵: 21:00-05:00

    早餐 --> [*]
    午餐 --> [*]
    晚餐 --> [*]
    夜宵 --> [*]
```

## 6. 报销生成流程

```mermaid
flowchart LR
    A[输入日期范围<br/>yymmdd-yymmdd] --> B[计算天数]

    B --> C[生成餐补记录<br/>40元×天数]
    B --> D[生成交补记录<br/>20元×天数]
    B --> E[生成住宿记录<br/>140元×天数]

    C --> F[合并到总表]
    D --> F
    E --> F

    F --> G[添加标签<br/>#差旅报销]
    G --> H[设置对象=天之逸]
    H --> I[输出 CSV]

    style A fill:#e1bee7
    style I fill:#c8e6c9
```

## 7. Git 自动同步流程

```mermaid
sequenceDiagram
    participant U as 用户
    participant A as Auto.py
    participant G as Git

    U->>A: python Auto.py -a
    A->>G: git add .
    A->>G: git commit -m "Auto-sync"
    A->>G: git pull --rebase

    alt 有冲突
        G-->>A: 冲突提示
        A->>U: 等待手动解决
    else 无冲突
        G-->>A: 拉取成功
        A->>G: git push
        G-->>U: 同步完成
    end
```

## 8. 配置数据优先级

```mermaid
graph TB
    A[INGREDIENT_PRIORITY] --> B[顺序很重要!]

    B --> C[1. 手动添加行<br/>最高优先级]
    C --> D[2. DATA_SOURCE<br/>商家规则匹配]
    D --> E[3. RAW_MAPPING_CONFIG<br/>兜底映射]

    style A fill:#ff6b6b
    style C fill:#4ecdc4
    style D fill:#ffe66d
    style E fill:#95e1d3

    F[⚠️ 修改禁忌] --> G[不要打乱顺序]
    F --> H[不要删除已有映射]
    F --> I[不要引用不存在的key]
```

## 使用说明

在支持 Mermaid 的 Markdown 预览器中打开本文件即可查看图表：

- **VS Code**: 安装 `Markdown Preview Mermaid Support` 插件
- **Typora**: 原生支持
- **GitHub**: 原生支持
- **在线工具**: https://mermaid.live

