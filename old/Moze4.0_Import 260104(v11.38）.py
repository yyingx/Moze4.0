"""
Moze 导入脚本 v11.39 (Prefix Fix & Auto-Close)
Created on Sun Jan 05 2026
@author: TZY_YX
"""

# ==========================================
#      配置与常量定义
# ==========================================

import datetime
import re
import traceback
import tkinter as tk
import time  # 引入 time 模块用于计时和延时
from tkinter import filedialog, simpledialog
from pathlib import Path
import numpy as np
import pandas as pd


class BColors:
    """终端输出颜色代码"""
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'


# --- 路径设置 ---
# 自动检测当前脚本所在目录
CURRENT_DIR = Path(__file__).parent if '__file__' in locals() else Path.cwd()

# 字典文件路径检测优先顺序
POSSIBLE_PATHS = [
    CURRENT_DIR / "Moze Dict.xlsx",
    Path(r"E:\天之逸2025\Moze4.0\Moze Dict.xlsx")
]

RULE_BOOK_PATH = next(
    (p for p in POSSIBLE_PATHS if p.exists()), CURRENT_DIR / "Moze Dict.xlsx")
TARGET_DIR = CURRENT_DIR / "Moze4.0_Import"

# --- 核心配置 ---
CONFIG = {
    'TRANSFER_TARGET_1': '肖恩',
    'TRANSFER_TARGET_2': '工商银行(9579)',
    'KEYWORD_CHARGING': '自助服务-充电桩',
    'KEYWORD_ALIPAY_CSV': '支付宝支付科技有限公司',
    'KEYWORD_WECHAT_XLSX': '微信支付账单明细',
}

# --- 账户标准化映射 (根据你的常用账户简化) ---
STANDARDIZE_ACCOUNTS = {
    r'.*4946.*': '平安银行4946',
    r'.*9579.*': '工商银行',
    r'.*3379.*': '招商银行Ⅱ',
    r'.*4826.*': '广发银行4826',
    r'.*零钱.*': '零钱3'
}

# --- 数据源定义 ---
DATA_SOURCE = {
    'FRUIT': [
        "水果",
        "果园", "果蔬", "百果园", "鲜果", "果业", "苹果", "香蕉", "红枣", "大枣", "桃子", "水蜜桃",
        "樱桃", "黄桃", "香梨", "雪梨", "柑橘", "沃柑", "草莓", "甘蔗", "桔子", "沙糖桔", "葡萄",
        "西瓜", "柚子", "柠檬", "橙子", "菠萝", "榴莲", "山竹", "蓝莓", "荔枝", "龙眼", "哈密瓜",
        "椰子", "柿子", "杨梅", "李子", "芒果", "猕猴桃", "火龙果", "百香果", "莲雾", "车厘子",
        "菠萝蜜", "桑葚", "枇杷", "杨桃", "无花果", "圣女果", "小番茄", "姑娘果",
        "西梅", "青枣", "冬枣", "黑布林", "人参果", "丑八怪", "耙耙柑"
    ],
    'DRINK': [
        "饮料",
        "可乐", "红牛", "奶茶", "东鹏", "果汁", "椰汁", "酸奶", "咖啡", "拿铁", "乐虎", "AD钙奶", "甜酒汁"
    ],
    'WATER': [
        "纯净水",
        "矿泉水", "农夫山泉", "怡宝", "百岁山", "娃哈哈", "今麦郎"
    ],
    'SOFTWARE': ["软件", "APP", "应用", "安卓"],
    'CHARGING': [
        "特来电", "星星充电", "小桔充电", "国家电网", "电费", "充电",
        "e充电", "云快充", "蔚来", "特斯拉", "超充"
    ],
    'VEGETABLE': [
        "蔬菜",
        # 叶菜类
        "白菜", "菠菜", "上海青", "油麦菜", "生菜", "娃娃菜", "空心菜", "苋菜", "菜芯",
        "韭菜", "香菜", "芹菜", "荠菜", "芥蓝",
        # 根茎类
        "土豆", "胡萝卜", "白萝卜", "青萝卜", "红薯", "紫薯", "山药", "芋头", "莲藕", "藕",
        "洋葱", "大蒜", "生姜", "竹笋", "芦笋", "茭白",
        # 茄果/瓜果类
        "西红柿", "茄子", "黄瓜", "西葫芦", "南瓜", "冬瓜", "苦瓜", "丝瓜", "青椒", "彩椒",
        "尖椒", "秋葵", "玉米",
        # 豆类
        "四季豆", "豇豆", "扁豆", "荷兰豆", "毛豆", "豌豆", "蚕豆", "刀豆",
        # 菌菇类
        "香菇", "金针菇", "平菇", "杏鲍菇", "口蘑", "木耳", "银耳", "茶树菇", "猴头菇",
        # 葱姜蒜/调味菜
        "大葱", "小葱", "蒜苗", "蒜苔"
    ],
    'Snack': [
        "零食",
        "切糕", "蛋糕", "面包", "腰果"
    ],
    'BEAN_PRODUCT': [
        "豆腐", "豆皮", "腐竹", "豆干", "豆腐泡", "千张", "豆卷", "腐皮", "油豆皮", "内脂豆腐",
        "老豆腐", "嫩豆腐", "冻豆腐", "响铃卷", "素鸡"
    ],
    'INGREDIENTS': [
        "食材",
        "馒头", "火腿", "腊肠", "榨菜", "老干妈", "杂酱", "生水饺", "酱料", "炸酱", "白砂糖",
        "蒸肉粉", "火锅底料", "食用盐", "生抽", "老抽", "耗油", "料酒", "胡椒粉", "辣椒粉",
        "辣椒酱", "鲜面条", "干面条", "挂面", "红豆", "绿豆", "黄豆"
    ],
    'PORK': [
        "猪肉", "扇子骨", "板油", "五花肉", "梅花肉", "瘦肉", "猪蹄膀", "蹄膀", "排骨", "脊骨",
        "五花", "前蹄", "大肠", "筒子骨", "棒骨", "猪油", "猪蹄", "荤油", "鲜肉"
    ],
    'POULTRY': ["三黄鸡", "鸡腿", "鸡翅", "鸡胸肉", "老母鸡", "乌鸡", "鸭肉", "鸭腿", "鸭翅", "鸭架", "老鸭"],
    'BEEF_MUTTON': ["牛肉", "牛腩", "牛排", "牛柳", "牛腱", "肥牛", "羊肉", "羊排", "羊腿"],
    'SEAFOOD': [
        "鱼", "虾", "蟹", "贝", "带鱼", "黄鱼", "鲈鱼", "鲫鱼", "草鱼", "基围虾", "皮皮虾",
        "大闸蟹", "生蚝", "鱿鱼", "章鱼", "海带", "紫菜"
    ],
    'Eggs': ["鸡蛋", "土鸡蛋", "生咸鸭蛋", "皮蛋"],
    'COOKED': [
        "熟食", "水煮花生", "花生米", "毛豆", "藕夹", "茄盒", "锅包肉", "猪头肉", "卤菜", "凉菜", "烧鸡",
        "烤鸭", "酱牛肉", "熏鱼", "炸带鱼", "炸鱿鱼", "肉丸", "墨鱼丸", "牛肉丸", "鱼丸", "虾滑",
        "贡丸", "鱿鱼圈", "糍粑", "熟咸鸭蛋", "熟肉肠", "糯米制品", "铁板鸭"
    ],
    'RICE': ["大米", "五常"],
    'MEAL': [
        "正餐",
        # 1. 地点/校区
        "东苑一层", "东苑二层", "西区食堂", "外勤",
        # 2. 连锁品牌
        "三镇民生", "永和四喜", "老乡鸡", "黄蜀郎", "麦香园", "丝路",
        # 3. 强特征的风味/地域
        "兰州", "沙县", "长沙臭豆腐", "重庆小面",
        # 4. 具体餐品 - 饭/面/粉/锅
        "麻辣香锅", "麻辣烫", "鸡公煲", "猪脚饭", "卤肉饭", "盖浇饭", "炒饭", "快餐", "小碗菜",
        "热干面", "太和板面", "板面", "油泼面", "牛肉面", "刀削面", "牛杂粉", "肠粉", "牛肉汤", "汤粉",
        # 5. 具体餐品 - 面点/早餐/小吃
        "水煎包", "小笼包", "汽水包", "煎豆折", "鸡蛋饼", "包粑",
        "煎包", "煎饼", "烧饼", "锅盔", "肉夹馍", "馕",
        "水饺", "蒸饺", "混沌", "馄饨", "包子", "小面",
        # 6. 通用场景/店名
        "烧烤", "路边摊", "食堂", "餐厅", "早点", "小吃", "餐饮", "面馆"
    ],
    'Parking_fee': ["WF7023"],
    'DAILY_NECESSITIES': [
        "日用",
        "抽纸", "卷纸", "厨房纸", "垃圾袋", "保鲜袋", "保鲜膜", "洗衣液", "洗洁精", "牙膏",
        "洗发水", "一次性手套", "一次性杯", "棉签"
    ],
    'Clothing_Shoes_Bags': ["袜子", "内裤", "帽子", "手套", "鞋", "T恤", "裤", "外套", "修裤脚"],
    'Adult_Products': ["避孕套", "成人润滑剂", "安全套"],
    'SERVER': ["节点", "Dler", "Dogess"],
    'Furniture_HomeTextiles': ["被子", "空调被", "枕头", "浴巾", "床笠"]
}

INGREDIENT_PRIORITY = [
    ('VEGETABLE', '蔬菜', '食材'),
    ('RICE', '大米', '食材'),
    ('BEAN_PRODUCT', '豆制品', '食材'),
    ('PORK', '猪肉', '食材'),
    ('BEEF_MUTTON', '牛羊肉', '食材'),
    ('POULTRY', '禽肉', '食材'),
    ('SEAFOOD', '海鲜水产', '食材'),
    ('Eggs', '蛋及蛋制品', '食材'),
    ('INGREDIENTS', '', '食材'),
    ('COOKED', '熟食', '食材'),
    ('Snack', '', '零食'),
    ('DAILY_NECESSITIES', '', '日常用品'),
    ('Clothing_Shoes_Bags', '', '服饰鞋包'),
    ('Furniture_HomeTextiles', '', '家具家纺'),
    ('SERVER', '节点', '虚拟其他'),
    ('Adult_Products', 'Condoms', '保健用品')
]

# (记录类型, 主类, 项目)
RAW_MAPPING_CONFIG = {
    ('支出', '饮食', '食'): ['午餐', '夜宵', '早餐', '晚餐', '纯净水', '零食', '食材', '饮料水果'],
    ('支出', '购物', '日用&家用'): ['家具家纺', '数码电器', '日常用品', '服饰鞋包', '大件'],
    ('支出', '交通', '通信&交通'): ['公共交通', '共享交通', '出租车', '汽车', '火车', '加油充电'],
    ('支出', '居家', '日用&家用'): ['快递邮政', '物业费', '理发'],
    ('支出', '居家', '通信&交通'): ['电话费'],
    ('支出', '居家', '住'): ['房租', '水费', '电费'],
    ('支出', '居家', '食'): ['液化气费'],
    ('支出', '医疗', '日用&家用'): ['体检', '药品', '门诊'],
    ('支出', '医疗', 'Hormones'): ['保健用品'],
    ('支出', '娱乐', 'Hormones'): ['休闲保健', '住宿'],
    ('支出', '娱乐', '娱乐'): ['旅游度假', '电影', '网游电玩'],
    ('支出', '学习', '学习'): ['图书', '证书'],
    ('支出', '个人', 'Hormones'): ['The Girls'],
    ('支出', '个人', '工作'): ['个人其他', '保险'],
    ('支出', '个人', '兼职'): ['生意'],
    ('支出', '个人', '额外非必要开销'): ['孝敬', '礼金红包', '社交人情', '给予'],
    ('支出', '虚拟', '娱乐'): ['App', '虚拟其他', '订阅'],
    ('支出', '虚拟', '学习'): ['Software'],
    ('收入', '收入', '兼职'): ['二手折旧', '外卖跑腿(CNY)'],
    ('收入', '收入', '工作'): ['收入其他', '福利补贴', '薪资'],
    ('收入', '收入', '理财'): ['利息'],
    ('收入', '收入', ''): ['红包'],
    ('应收款项', '应收款项', ''): ['报账', '借出', '代付', '押金'],
    ('应付款项', '应付款项', ''): ['借入']
}

PATTERNS = {k: r"(?:" + "|".join(map(re.escape, v)) +
            ")" for k, v in DATA_SOURCE.items()}

AUTO_MAP_DICT = {}
for (r_type, main_cat, proj), sub_list in RAW_MAPPING_CONFIG.items():
    for sub in sub_list:
        AUTO_MAP_DICT[sub] = (r_type, main_cat, proj)

COLUMN_MAPPING = {
    '交易时间': '交易时间',
    '交易创建时间': '交易时间',
    '付款时间': '交易时间',
    '交易类型': '交易类型',
    '交易分类': '交易类型',
    '交易对方': '交易对方',
    '商品': '商品',
    '商品名称': '商品',
    '收/支': '收/支',
    '金额(元)': '金额(元)',
    '金额（元）': '金额(元)',
    '金额': '金额(元)',
    '支付方式': '支付方式',
    '收/付款方式': '支付方式',
    '当前状态': '当前状态',
    '交易状态': '当前状态',
    '备注': '备注',
    '交易单号': '交易单号',
    '交易订单号': '交易单号',
    '商户单号': '商户单号',
    '商家订单号': '商户单号',
    '商品说明': '商品'
}

# ==========================================
#          [核心] 逻辑处理函数
# ==========================================


def construct_description(df):
    """构建初始描述：备注 > 字典规则 > 淘宝商品名"""
    df['描述'] = ""
    mask_tb = df.get('商家', '') == '淘宝'
    if mask_tb.any():
        df.loc[mask_tb, '描述'] = df.loc[mask_tb, '商品'].fillna("").astype(str)

    if '备注' in df.columns:
        raw_memo = df['备注'].astype(str).replace('nan', '').str.strip()
        df['描述'] = np.where(raw_memo != "", raw_memo, df['描述'])

    if '描述_rule' in df.columns:
        df['描述'] = df['描述_rule'].combine_first(df['描述'])

    df['描述'] = df['描述'].astype(str).str.replace(
        r'[\r\n]+', ' ', regex=True).str.strip()
    df.loc[df['描述'] == '/', '描述'] = ""
    return df


def clean_final_description(df):
    """
    最终清洗：去前缀 & 去冗余词
    范围扩大：包含所有[子类] + [Name名称] + [Meal餐饮关键词]
    """
    
    # 1. --- 聚合所有触发词 (Trigger Words) ---
    triggers = set()

    # (A) 所有子类 (Sub-categories)
    triggers.update(AUTO_MAP_DICT.keys())

    # (B) 所有餐饮关键词 (Meal)
    # if 'MEAL' in DATA_SOURCE:
    #     triggers.update(DATA_SOURCE['MEAL'])

    # (C) 所有名称 (Name)
    # 来源 1：从 INGREDIENT_PRIORITY 提取
    triggers.update([x[1] for x in INGREDIENT_PRIORITY if x[1]])
    
    # 来源 2：常见手动名称
    manual_names = [
        "水果", "饮料", "纯净水", "充电", "加油充电", 
        "Software", "停车费", "日用", "正餐", "零食"
    ]
    triggers.update(manual_names)

    # 2. --- 预处理列表 ---
    # 移除空字符串和纯数字
    trigger_list = [w for w in triggers if w and not str(w).isdigit()]
    
    # 【核心】按长度倒序排列，防止短词误伤长词
    trigger_list.sort(key=len, reverse=True)

    # 3. --- 执行清洗 ---
    
    # [操作 A] 去前缀
    keywords_regex = "|".join(map(re.escape, trigger_list))
    prefix_pat = rf"^(?:{keywords_regex})[. 。\s-]+(.*)"
    
    df['描述'] = df['描述'].astype(str).str.replace(
        prefix_pat, r'\1', regex=True).str.strip()

    # [操作 B] 去完全冗余
    mask_redundant = df['描述'].isin(trigger_list) | (df['描述'] == "")
    if mask_redundant.any():
        df.loc[mask_redundant, '描述'] = ""

    return df


def robust_date_converter(x):
    """健壮的日期转换"""
    if isinstance(x, (datetime.datetime, datetime.date)):
        return pd.to_datetime(x)
    try:
        if isinstance(x, (int, float)):
            return pd.to_datetime(x, unit='D', origin='1899-12-30')
        return pd.to_datetime(x, errors='coerce')
    except Exception:
        return pd.NaT


def load_settings(rule_path: Path):
    """加载配置文件"""
    if not rule_path.exists():
        return
    try:
        df = pd.read_excel(rule_path, sheet_name='Settings', engine='openpyxl')
        for _, row in df.iterrows():
            k, v = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
            if k in CONFIG and v and v.lower() != 'nan':
                CONFIG[k] = v
        print("已加载自定义配置")
    except Exception:
        pass


def load_rules(rule_path: Path):
    """加载规则"""
    print("正在加载商家映射规则...", flush=True)
    try:
        xl = pd.ExcelFile(rule_path, engine='openpyxl')
        s_name = next(
            (s for s in xl.sheet_names
             if s.lower() != 'settings' and
             '商家(old)' in pd.read_excel(rule_path, sheet_name=s, nrows=1).columns),
            None
        )
        if not s_name:
            return None
        df = pd.read_excel(rule_path, sheet_name=s_name, engine='openpyxl')
        df['is_regex'] = pd.to_numeric(
            df.get('is_regex', 0), errors='coerce').fillna(0)
        df['商家(old)'] = df['商家(old)'].astype(str).str.strip()
        return df
    except Exception:
        return None


def sniff_and_load_data(file_path: Path):
    """嗅探并加载数据"""
    print(f"读取: {file_path.name}", flush=True)
    ftype, df = None, None

    try:
        # 1. 支付宝 (CSV)
        if file_path.suffix.lower() == '.csv':
            for enc in ['gb18030', 'utf-8']:
                try:
                    preview = pd.read_csv(
                        file_path, header=None, nrows=50, encoding=enc, names=list(range(30))
                    )
                    mask = preview.apply(
                        lambda x: x.astype(str).str.contains(
                            '交易创建时间|交易时间|付款时间', na=False
                        )
                    ).any(axis=1)

                    if mask.any():
                        header_row = mask.idxmax()
                    else:
                        header_row = 24

                    df = pd.read_csv(
                        file_path, header=header_row, encoding=enc)
                    ftype = "AliPay"
                    print(f"  (支付宝: {enc}, 表头锁定第 {header_row+1} 行)")
                    break
                except (UnicodeDecodeError, pd.errors.ParserError):
                    continue

            if df is None:
                print("  无法读取支付宝文件 (编码或格式无法识别)")
                return None

        # 2. 微信 (Excel)
        elif file_path.suffix.lower() == '.xlsx':
            try:
                df = pd.read_excel(file_path, header=16, engine='openpyxl')
                ftype = "WeChat"
            except Exception as e:
                print(f"  微信Excel读取错误: {e}")
                return None

        if df is None:
            return None

        df.columns = df.columns.astype(str).str.strip().str.replace(
            '（', '(', regex=False).str.replace('）', ')', regex=False
                                               )
        df.rename(columns=COLUMN_MAPPING, inplace=True)

        if '金额(元)' not in df.columns:
            print(f"  失败: 未找到'金额(元)'列。当前列名: {df.columns.tolist()}")
            return None

        df['金额(元)'] = pd.to_numeric(
            df['金额(元)'].astype(str).str.replace('¥', '', regex=False)
            .str.replace(',', '', regex=False).str.strip(),
            errors='coerce'
        ).fillna(0)

        df['_source_tag'] = '#WechatPay' if ftype == "WeChat" else '#AliPay'
        print(f"  成功读取 {len(df)} 行 (类型: {ftype})。")
        return df

    except Exception as e:
        print(f"  加载未知错误: {e}")
        return None


def ensure_columns(df, main_col, sub_col):
    """确保必要的列存在"""
    cols_to_check = [
        main_col, sub_col, '金额', '记录类型', '账户', '商家', '描述',
        '项目', '名称', '对象', '标签', '日期', '时间', '币种', '手续费', '折扣'
    ]
    for c in cols_to_check:
        if c not in df.columns:
            df[c] = pd.NA

    cols_to_fill = ['记录类型', '项目', '对象', main_col, sub_col, '商家', '名称']
    for c in cols_to_fill:
        df[c] = df[c].fillna("")
    return df


def process_transfers(df, main_col, sub_col):
    """处理转账 (双向生成)"""
    res = []
    t1, t2 = CONFIG['TRANSFER_TARGET_1'], CONFIG['TRANSFER_TARGET_2']

    def add_pair(sub_df, out_a, in_a, sub_val='转账'):
        if sub_df.empty:
            return
        o_rec, i_rec = sub_df.copy(), sub_df.copy()

        # 构建转出记录
        o_rec['金额'] = pd.to_numeric(o_rec['金额(元)']) * -1
        o_rec['记录类型'] = '转出'
        o_rec['账户'] = out_a
        o_rec[sub_col] = sub_val

        # 构建转入记录
        i_rec['金额'] = pd.to_numeric(i_rec['金额(元)'])
        i_rec['记录类型'] = '转入'
        i_rec['账户'] = in_a
        i_rec[sub_col] = sub_val

        res.extend([o_rec, i_rec])

    # 逻辑：Target1 视为零钱3与零钱2的互转
    add_pair(df[(df["交易对方"] == t1) & (df["收/支"] == "收入")], '零钱2', '零钱3')
    add_pair(df[(df["交易对方"] == t1) & (df["收/支"] == "支出")], '零钱3', '零钱2')

    # 逻辑：Target2 视为提现到工行
    add_pair(df[df["交易对方"] == t2], '零钱3', '工商银行', '提现')

    if not res:
        return pd.DataFrame()

    ret = pd.concat(res)
    ret = ensure_columns(ret, main_col, sub_col)
    ret[main_col] = "转账"
    ret['日期'] = ret['交易时间'].dt.strftime('%Y/%m/%d')
    ret['时间'] = ret['交易时间'].dt.strftime('%H:%M:%S')
    ret['_Sort_Date'] = ret['交易时间']
    ret.loc[:, ['币种', '手续费', '折扣']] = ["CNY", 0, 0]
    return ret


def process_heuristics(df, main_col, sub_col):
    """启发式处理"""
    df = ensure_columns(df, main_col, sub_col)
    for col in [sub_col, '项目', '描述', '商家', '商品']:
        df[col] = df[col].astype(str).str.strip().replace("nan", "")

    def check(k):
        pat = PATTERNS.get(k, "^$")
        return (
            df['商家'].str.contains(pat, regex=True) |
            df['商品'].str.contains(pat, regex=True) |
            df['描述'].str.contains(pat, regex=True)
        )

    uncat = df[sub_col] == ""

    # 1. 关键词分类
    mask = df['商品'].str.contains(CONFIG['KEYWORD_CHARGING'], na=False) & uncat
    if mask.any():
        df.loc[mask, ['名称', sub_col]] = ['充电', '加油充电']

    mask = check('SOFTWARE') & uncat
    if mask.any():
        df.loc[mask, sub_col] = 'Software'

    for k, n, s in [('FRUIT', '水果', '饮料水果'), ('DRINK', '饮料', '饮料水果'), ('WATER', '', '纯净水')]:
        mask = check(k) & uncat & ((df[sub_col] == "") | (df[sub_col] == "食材"))
        if mask.any():
            df.loc[mask, ['名称', sub_col]] = [n, s]

    mask_meal = check('MEAL')
    for key, name, sub_c in INGREDIENT_PRIORITY:
        mask = check(key) & (~mask_meal) & (
            uncat | (df[main_col].isin(['购物', '居家', '饮食'])))
        if key in ['SEAFOOD', 'PORK', 'POULTRY', 'BEEF_MUTTON', 'VEGETABLE']:
            mask &= (~check('COOKED'))
        if mask.any():
            df.loc[mask, [sub_col, '名称']] = [sub_c, name]

    # 停车费
    mask = check('Parking_fee')
    if mask.any():
        df.loc[mask, [sub_col, '名称', '对象']] = ['报账', '停车费', '天之逸']

    # 2. 债权提取
    keys = ['报账', '借出', '代付', '押金', '借入']
    debt_pat = rf"({'|'.join(keys)})\s*(.*)"
    extracted = df['描述'].str.extract(debt_pat, expand=True)
    mask_found = extracted[0].notna()
    if mask_found.any():
        df.loc[mask_found, sub_col] = extracted[0]
        mask_name = mask_found & (extracted[1].str.strip() != "")
        if mask_name.any():
            df.loc[mask_name, '对象'] = extracted[1][mask_name].str.strip()
        df.loc[mask_found, ['项目', '描述']] = ""

    # ================== 【升级修复：前缀提取分类】 ==================
    # 逻辑：不仅检查“描述 == 子类”，还检查“描述”是否以“子类 + 分隔符”开头
    # 例子：描述为 "大件.摩托车"，能识别出 "大件"； 描述为 "休闲保健"，能识别 "休闲保健"
    
    uncat = df[sub_col] == ""
    
    # 1. 准备关键词 (按长度倒序，防止短词误判，如 "饮料水果" 优于 "饮料")
    valid_subcats = sorted(list(AUTO_MAP_DICT.keys()), key=len, reverse=True)
    
    # 2. 构建正则：^(子类1|子类2|...)(?:[. 。\s]|$|-)
    # 解释：
    #   ^       : 必须出现在开头
    #   (...)   : 捕获组，提取出具体的子类名称
    #   (?:...) : 非捕获组，后面必须跟着 [点/空格/句号/横杠] 或者 [字符串结束]
    #   这样防止 "大件物品" 被错误识别为 "大件" (如果没有分隔符的话)
    pattern_prefix = rf"^({'|'.join(map(re.escape, valid_subcats))})(?:[. 。\s]|$|-)"
    
    # 3. 提取匹配
    # 只对未分类的行进行提取
    if not uncat.empty and valid_subcats:
        extracted_series = df.loc[uncat, '描述'].astype(str).str.extract(pattern_prefix, expand=False)
        
        # 4. 填入结果
        mask_hit = extracted_series.notna()
        if mask_hit.any():
            # 找到原本 df 中的索引
            target_indices = uncat[uncat].index[mask_hit]
            print(f"  [提示] 发现 {len(target_indices)} 条记录通过【前缀提取】匹配子类 (如: {extracted_series[mask_hit].iloc[0]}.xx)")
            df.loc[target_indices, sub_col] = extracted_series[mask_hit]
            
    # ================== 【升级修复 结束】 ==================

    # 3. [核心] 全自动推导 (子类 -> 记录类型, 主类, 项目)
    mapped_values = df[sub_col].map(AUTO_MAP_DICT)
    mask_mapped = mapped_values.notna()
    if mask_mapped.any():
        target_types = mapped_values[mask_mapped].apply(lambda x: x[0])
        target_mains = mapped_values[mask_mapped].apply(lambda x: x[1])
        target_projs = mapped_values[mask_mapped].apply(lambda x: x[2])

        df.loc[mask_mapped, '记录类型'] = target_types
        df.loc[mask_mapped, main_col] = target_mains
        df.loc[mask_mapped, '项目'] = target_projs

    return df


def process_main(df, df_rules, main_col, sub_col):
    """主处理逻辑：规则匹配、分类、清洗"""
    # 初始化缺失列
    for c in ['当前状态', '收/支', '交易对方', '交易时间']:
        if c not in df.columns:
            df[c] = ""

    # 预筛选：排除退款和关闭的交易，只处理支出（转账已在别处处理）
    df = df[
        (~df["当前状态"].isin(["已全额退款", "交易关闭"])) &
        (df["收/支"] == "支出")
    ].copy()

    if df.empty:
        return pd.DataFrame()
    print(f"--- 处理主交易 ({len(df)}条) ---", flush=True)

    df = ensure_columns(df, main_col, sub_col)

    # 1. 支付方式标准化 (使用顶部定义的字典)
    df['支付方式'] = df['支付方式'].astype(str).replace(
        STANDARDIZE_ACCOUNTS, regex=True)

    # 2. 规则匹配 (Merge方式)
    df['商家(old)'] = df.get('交易对方', pd.NA).astype(str).str.strip()
    rename_map = {c: f'{c}_rule' for c in df_rules.columns if c not in [
        '商家(old)', 'is_regex']}

    # 精确匹配
    df = pd.merge(
        df,
        df_rules[df_rules['is_regex'] == 0]
        .rename(columns=rename_map)
        .drop(columns='is_regex', errors='ignore'),
        on='商家(old)', how='left'
    )

    # 正则匹配 (仅对未匹配项)
    uncat = df[df[f"{main_col}_rule"].isna()].index if f"{
        main_col}_rule" in df.columns else df.index
    if not uncat.empty:
        to_match = df.loc[uncat, '交易对方'].astype(str)
        for _, r in df_rules[df_rules['is_regex'] == 1].iterrows():
            idx = to_match[to_match.str.contains(
                r['商家(old)'], regex=True)].index
            if not idx.empty:
                for c, rc in rename_map.items():
                    df.loc[idx, rc] = r[c]
                to_match = to_match.drop(idx)

    # 应用规则覆盖
    for orig, rule in rename_map.items():
        if rule in df.columns:
            if orig != '描述':
                df[orig] = df[rule].combine_first(df[orig])
            if orig != '描述':
                df.drop(columns=[rule], inplace=True)

    # 3. 淘宝特殊识别
    if '商户单号' in df.columns:
        mask_tb = (df.get('_source_tag', '') == '#AliPay') & \
            df['商户单号'].astype(str).str.strip(
        ).str.startswith('T200P4', na=False)
        if mask_tb.any():
            df.loc[mask_tb, '商家'] = '淘宝'

    # 4. 构建描述
    df = construct_description(df)
    if '描述_rule' in df.columns:
        df.drop(columns=['描述_rule'], inplace=True)

    df['记录类型'] = df.get('记录类型', pd.NA).replace("", pd.NA).fillna(df['收/支'])
    df['商家'] = df.get('商家', pd.NA).replace("", pd.NA).fillna(df['交易对方'])

    # 5. 餐饮时段细化
    df[sub_col] = df[sub_col].astype(str).str.strip().replace("nan", "")
    mask_time = (df[sub_col] == "") & (
        df['商家(old)'].str.contains(PATTERNS['MEAL'], regex=True, na=False) |
        df['商家'].str.contains(PATTERNS['MEAL'], regex=True, na=False) |
        df['商品'].str.contains(PATTERNS['MEAL'], regex=True, na=False) |
        df['描述'].str.contains(PATTERNS['MEAL'], regex=True, na=False)
    )
    if mask_time.any():
        h = df.loc[mask_time, '交易时间'].dt.hour
        df.loc[mask_time, sub_col] = np.select(
            [(h >= 6) & (h < 11), (h >= 11) & (h < 16), (h >= 16) & (h < 21)],
            ["早餐", "午餐", "晚餐"],
            default="夜宵"
        )

    # 6. 启发式分类 & 全自动推导
    df = process_heuristics(df, main_col, sub_col)

    # 7. 最终数据清洗
    df['金额'] = pd.to_numeric(df.get('金额(元)', 0), errors='coerce').fillna(0)
    # 支出转为负数
    df.loc[
        (df['记录类型'].isin(['支出', '应付款项', '应收款项'])) &
        (df.get('收/支') == '支出'), '金额'
    ] *= -1

    # 修正借入金额符号和支付方式
    df.loc[df[sub_col] == '借入', '金额'] = df.loc[df[sub_col] == '借入', '金额'].abs()
    df.loc[(df['支付方式'] == "/") & (df['记录类型'] == "收入"), '支付方式'] = "零钱3"

    df['日期'] = df['交易时间'].dt.strftime('%Y/%m/%d')
    df['时间'] = df['交易时间'].dt.strftime('%H:%M:%S')
    df['_Sort_Date'] = df['交易时间']
    df['账户'] = df.get('支付方式', "")
    df.loc[:, ['币种', '手续费', '折扣']] = ["CNY", 0, 0]

    # 清理非商家字段
    df.loc[df['记录类型'].isin(['应收款项', '应付款项']), '商家'] = ""
    mask_debt = df['记录类型'].isin(['应收款项', '应付款项'])
    if mask_debt.any():
        df.loc[mask_debt, '项目'] = ""

    df = clean_final_description(df)

    return df


def get_user_input():
    """获取用户输入文件和日期（修复 Spyder 弹窗焦点问题）"""
    # 创建隐藏的主窗口
    root = tk.Tk()
    root.withdraw()

    # --- 核心：夺取焦点五连鞭 ---
    # 1. 设置置顶
    root.attributes('-topmost', True)
    # 2. 提升层级
    root.lift()
    # 3. 强制获取焦点
    root.focus_force()
    # 4. 强制刷新事件循环 (确保系统执行上述指令)
    root.update()

    print("\n选择文件...")

    # 5. 绑定父窗口 (parent=root)，让弹窗继承置顶属性
    files = filedialog.askopenfilenames(
        parent=root,
        title="请选择 Moze 导出的文件",
        filetypes=[("Excel/CSV", "*.xlsx;*.csv")]
    )

    # 选完文件后取消置顶，防止遮挡后续操作
    root.attributes('-topmost', False)
    root.update()

    if not files:
        root.destroy()
        return None, None

    # 日期输入框同样需要绑定 parent
    start = simpledialog.askstring(
        "筛选",
        "起始日期 yymmdd (留空导入全部):",
        parent=root
    )
    st_date = pd.to_datetime(start, format="%y%m%d") if start else None

    # 彻底销毁窗口
    root.destroy()
    return files, st_date


def save_result(df, cols):
    """保存结果文件"""
    # 标签处理
    if '标签' in df.columns:
        m = ~df['记录类型'].isin(['转入', '转出'])
        df.loc[m, '标签'] = (
            df.loc[m, '标签'].fillna("") + " " + df.loc[m, '_source_tag']
        ).str.strip()

    # 排序：时间(新->旧), 优先级(0->1)
    type_priority = {'转出': 0, '转入': 1}
    df['_Type_Rank'] = df['记录类型'].map(type_priority).fillna(0)
    df.sort_values(
        ['_Sort_Date', '_Type_Rank'], ascending=[False, True], inplace=True
    )

    if not TARGET_DIR.exists():
        TARGET_DIR.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    path = TARGET_DIR / f"MOZE导入_{timestamp}.csv"

    df.to_csv(path, index=False, columns=cols, encoding='utf-8-sig')
    return path


def main():
    """主程序入口"""
    print(f"{BColors.BOLD}=== Moze 导入工具 v11.39 (Prefix Fix & Auto-Close) ==={BColors.ENDC}")
    try:
        load_settings(RULE_BOOK_PATH)
        df_rules = load_rules(RULE_BOOK_PATH)
        if df_rules is None:
            input("按 Enter 退出")
            return

        cols = [c for c in df_rules.columns if c not in [
            '商家(old)', 'is_regex']]
        main_col = next((c for c in cols if "主类" in c), None)
        sub_col = next((c for c in cols if "子类" in c), None)

        if not main_col:
            print("错：找不到主类列")
            return

        files, st_date = get_user_input()
        if not files:
            return

        dfs = [d for f in files if (
            d := sniff_and_load_data(Path(f))) is not None]
        if not dfs:
            return

        df_raw = pd.concat(dfs, ignore_index=True)
        df_raw['交易时间'] = df_raw['交易时间'].apply(robust_date_converter)

        if st_date:
            df_raw = df_raw[df_raw['交易时间'] >= st_date]

        df_raw['交易对方'] = df_raw.get('交易对方', pd.NA).astype(str).str.strip()

        # 分流处理：转账 vs 普通交易
        mask_trans = df_raw['交易对方'].isin([
            CONFIG['TRANSFER_TARGET_1'], CONFIG['TRANSFER_TARGET_2']
        ])

        res_dfs = []
        if mask_trans.any():
            res_dfs.append(process_transfers(
                df_raw[mask_trans].copy(), main_col, sub_col))
        if (~mask_trans).any():
            res_dfs.append(process_main(
                df_raw[~mask_trans].copy(), df_rules, main_col, sub_col))

        if not res_dfs:
            print("无结果")
            return

        df_final = pd.concat(res_dfs, ignore_index=True)

        # 补全所有列
        for c in cols:
            if c not in df_final.columns:
                df_final[c] = ""

        path = save_result(df_final, cols)
        print(f"\n{BColors.OKGREEN}成功! 文件: {path}{BColors.ENDC}")

        # 最终校验
        checks = [
            ('支出', df_final['记录类型'] == '支出', [main_col, sub_col, '项目']),
            ('债权', df_final['记录类型'].isin(
                ['应收款项', '应付款项']), [main_col, sub_col, '对象'])
        ]

        for name, mask, check_cols in checks:
            bad = mask & df_final[check_cols].isin(["", pd.NA]).any(axis=1)
            if bad.any():
                print(f"{BColors.FAIL}[严重] {bad.sum()} 条【{
                      name}】记录缺失关键信息:{BColors.ENDC}")
                print(
                    df_final.loc[bad, ['日期', '商家', '描述', '金额'] + check_cols]
                    .head().to_string()
                )
            else:
                print(f"{BColors.OKGREEN}√ {name}记录完美{BColors.ENDC}")

    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    start_time = time.time()  # 1. 开始计时
    
    main()
    
    end_time = time.time()    # 2. 结束计时
    duration = end_time - start_time
    
    print(f"\n{BColors.BOLD}--------------------------------{BColors.ENDC}")
    print(f"运行结束 | 总耗时: {duration:.2f} 秒")
    print(f"{BColors.BOLD}--------------------------------{BColors.ENDC}")

    # 3. 倒计时自动退出 (省去按回车)
    try:
        for i in range(3, 0, -1):
            print(f"\r程序将在 {i} 秒后自动关闭...", end="", flush=True)
            time.sleep(1)
    except KeyboardInterrupt:
        pass