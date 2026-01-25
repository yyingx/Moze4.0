# -*- coding: utf-8 -*-
"""
Moze 导入脚本 v11.66 (Taobao Fix & Memo Subcat)
Created on Sun Jan 05 2026
Optimized: Sat Jan 25 2026

@author: TZY_YX

BUG FIXES (v11.66):
[已修复] 淘宝订单商家未识别 → 商户单号前缀从 T200P4 改为 T200P（兼容 T200P4/T200P5 等）
[新增] 备注子类别解析 → 备注格式"子类别+描述"自动映射，如"探索claude"→子类=探索,描述=claude
[优化] 移除点语法处理，简化代码结构
"""

import numpy as np
import pandas as pd
from pathlib import Path
from tkinter import filedialog, simpledialog
import time
import tkinter as tk
import traceback
import re
import datetime

# 抑制 SettingWithCopyWarning
pd.options.mode.chained_assignment = None


class BColors:
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'


# --- 路径设置 ---
CURRENT_DIR = Path(__file__).parent if '__file__' in locals() else Path.cwd()
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
    'TRANSFER_TARGET_SNOWBALL': '上海雪球数智科技有限公司',
}

STANDARDIZE_ACCOUNTS = {
    r'.*4946.*': '平安银行4946',
    r'.*9579.*': '工商银行',
    r'.*3379.*': '招商银行Ⅱ',
    r'.*4826.*': '广发银行4826',
    r'.*零钱.*': '零钱3'
}

# --- 1. DATA_SOURCE ---
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
        "可乐", "红牛", "奶茶", "东鹏", "果汁", "椰汁", "酸奶", "咖啡", "拿铁", "乐虎", "AD钙奶",
        "蜜雪冰城"
    ],
    'WATER': [
        "纯净水",
        "矿泉水", "农夫山泉", "怡宝", "百岁山", "娃哈哈", "今麦郎"
    ],
    'VEGETABLE': [
        "蔬菜",
        # 叶菜类
        "白菜", "菠菜", "上海青", "油麦菜", "生菜", "娃娃菜", "空心菜", "苋菜", "菜芯",
        "韭菜", "香菜", "芹菜", "荠菜", "芥蓝",
        # 根茎类
        "土豆", "胡萝卜", "萝卜", "白萝卜", "青萝卜", "红薯", "紫薯", "山药", "芋头", "莲藕", "藕",
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
        # --- 通用 ---
        "食材",

        # --- 面点主食 ---
        "馒头", "发糕", "生水饺", "鲜面条", "干面条", "挂面",

        # --- 腊味腌货 ---
        "火腿", "腊肠", "榨菜", "甜酒",

        # --- 调味品/酱料 ---
        "老干妈", "杂酱", "酱料", "炸酱", "酱豆", "辣椒酱",
        "白砂糖", "食用盐", "生抽", "老抽", "耗油", "料酒",
        "胡椒粉", "辣椒粉", "蒸肉粉", "火锅底料",

        # --- 干货杂粮 ---
        "红豆", "绿豆", "黄豆"
    ],
    'PORK': [
        "猪肉", "扇子骨", "板油", "五花肉", "梅花肉", "瘦肉", "猪蹄膀", "蹄膀", "排骨", "脊骨",
        "五花", "前蹄", "大肠", "筒子骨", "棒骨", "猪油", "猪蹄", "荤油", "鲜肉"
    ],
    'POULTRY': ["三黄鸡", "鸡腿", "鸡翅", "鸡胸肉", "老母鸡", "乌鸡", "鸭肉", "鸭腿", "鸭翅", "鸭架", "老鸭"],
    'BEEF_MUTTON': ["牛肉", "牛腩", "牛排", "牛柳", "牛肠", "牛杂", "牛腱", "肥牛", "羊肉", "羊排", "羊腿"],
    'SEAFOOD': [
        "鱼", "虾", "蟹", "贝", "带鱼", "黄鱼", "鲈鱼", "鲫鱼", "草鱼", "基围虾", "皮皮虾",
        "大闸蟹", "生蚝", "鱿鱼", "章鱼", "海带", "紫菜"
    ],
    'Eggs': ["鸡蛋", "土鸡蛋", "生咸鸭蛋", "皮蛋"],
    'COOKED': [
        "熟食", "水煮花生", "花生米", "蚕豆", "毛豆", "藕夹", "茄盒", "锅包肉", "猪头肉", "卤菜", "凉菜", "烧鸡",
        "烤鸭", "酱牛肉", "熏鱼", "炸带鱼", "炸鱿鱼", "肉丸", "墨鱼丸", "牛肉丸", "鱼丸", "虾滑",
        "贡丸", "鱿鱼圈", "糍粑", "熟咸鸭蛋", "熟肉肠", "糯米制品", "铁板鸭"
    ],
    'RICE': ["大米", "五常"],
    'MEAL': [
        "正餐",
        # 1. 地点/校区
        "东苑一层", "东苑二层", "西区食堂", "外勤", "斯迪姆幼儿园-柏思思",
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
    'DAILY_NECESSITIES': [
        "日用",
        "抽纸", "卷纸", "厨房纸", "垃圾袋", "保鲜袋", "保鲜膜", "洗衣液", "洗洁精", "牙膏",
        "洗发水", "一次性手套", "一次性杯", "棉签", "纸巾"
    ],
    'Clothing_Shoes_Bags': ["袜子", "内裤", "帽子", "手套", "鞋", "T恤", "裤", "外套", "修裤脚"],
    'Adult_Products': ["避孕套", "成人润滑剂", "安全套", "Condoms"],
    'SOFTWARE': ["软件", "APP", "应用", "安卓"],
    'SERVER': ["节点", "Dler", "Dogess"],
    'Furniture_HomeTextiles': ["被子", "空调被", "枕头", "浴巾", "床笠"],
    'Parking_fee': ["WF7023"],
    'CHARGING': ["自助服务-充电桩"],  # 充电桩关键词
    'REIM_TRAVEL': [
        "车船费", "住宿费", "住宿补贴", "交通补贴", "餐费补贴"
    ],
    'REIM_EXPENSE': [
        "材料费", "燃油费", "交通费", "过路费", "租赁费",
        "叉车费", "停车费", "印刷服务", "物流运输", "市内交通",
        "生活用品", "人工劳务费",
        "代付货款", "招待费", "汽车费用", "代付"
    ]
}

INGREDIENT_PRIORITY = [
    ('DRINK', '饮料', '饮料水果'),
    ('FRUIT', '水果', '饮料水果'),
    ('WATER', '', '纯净水'),
    ('VEGETABLE', '蔬菜', '食材'),
    ('RICE', '大米', '食材'),
    ('BEAN_PRODUCT', '豆制品', '食材'),
    ('PORK', '猪肉', '食材'),
    ('BEEF_MUTTON', '牛羊肉', '食材'),
    ('POULTRY', '禽肉', '食材'),
    ('Eggs', '蛋及蛋制品', '食材'),
    ('SEAFOOD', '海鲜水产', '食材'),
    ('COOKED', '熟食', '食材'),
    ('INGREDIENTS', '', '食材'),
    ('Snack', '', '零食'),
    ('CHARGING', '充电', '加油充电'),
    ('SOFTWARE', '', 'Software'),
    ('SERVER', '节点', '虚拟其他'),
    ('DAILY_NECESSITIES', '', '日常用品'),
    ('Clothing_Shoes_Bags', '', '服饰鞋包'),
    ('Furniture_HomeTextiles', '', '家具家纺'),
    ('Adult_Products', 'Condoms', '保健用品')
]

RAW_MAPPING_CONFIG = {
    ('支出', '饮食', '食'): ['早餐', '午餐', '晚餐', '夜宵', '正餐', '食材', '饮料水果', '纯净水', '零食'],
    ('支出', '购物', '日用&家用'): ['服饰鞋包', '日常用品', '大件', '共享租赁', '摄影文印', '数码电器', '家具家纺'],
    ('支出', '交通', '通信&交通'): ['加油充电', '公共交通', '共享交通', '火车', '出租车', '汽车', '轮渡', '机票', '交通违章', '维修保养', '停车费'],
    ('支出', '居家', '住'): ['房租', '水费', '电费', '物业费', '宽带费'],
    ('支出', '居家', '食'): ['液化气费'],
    ('支出', '居家', '日用&家用'): ['快递邮政', '理发', '洗衣费'],
    ('支出', '居家', '通信&交通'): ['电话费'],
    ('支出', '虚拟', '娱乐'): ['App', '订阅', '虚拟其他', '影音'],
    ('支出', '虚拟', '学习'): ['Software'],
    ('支出', '娱乐', '娱乐'): ['电影', '聚会', '旅游度假', '卡拉OK', '麻将棋牌', '网游电玩', '娱乐其他'],
    ('支出', '娱乐', 'Hormones'): ['休闲保健', '住宿'],
    ('支出', '医疗', '医疗'): ['医疗用品', '牙齿保健', '药品', '门诊', '打针', '住院', '手术', '体检'],
    ('支出', '医疗', 'Hormones'): ['保健用品'],
    ('支出', '学习', '学习'): ['图书', '教材', '证书', '探索', '文具', '资料文献'],
    ('支出', '个人', 'Hormones'): ['The Girls'],
    ('支出', '个人', '工作'): ['个人其他', '保险'],
    ('支出', '个人', '兼职'): ['生意', '投资亏损'],
    ('支出', '个人', '理财'): ['利息'],
    ('支出', '个人', '额外非必要开销'): ['社交人情', '给予', '孝敬', '礼金红包'],
    ('收入', '收入', '兼职'): ['外卖跑腿(CNY)', '其他收入'],
    ('收入', '收入', '工作'): ['薪资', '福利补贴', '年终奖'],
    ('收入', '收入', '理财'): ['利息收入', '投资盈利'],
    ('收入', '收入', ''): ['收红包', '二手折旧'],
    ('转出', '转账', ''): ['转账', '提现', '取出', '存款', '兑换', '充值'],
    ('转入', '转账', ''): ['转账', '提现', '取出', '存款', '兑换', '充值'],
    ('转出', '信用卡还款', ''): ['信用卡还款'],
    ('转入', '信用卡还款', ''): ['信用卡还款'],
    ('应收款项', '应收款项', ''): ['报账', '借出', '代付', '押金'],
    ('应付款项', '应付款项', ''): ['借入'],
    ('手续费', '手续费', ''): ['手续费'],
    ('折扣', '折扣', ''): ['折扣'],
    ('返利回馈', '返利回馈', '返利'): ['返利回馈']
}

PATTERNS = {k: re.compile(r"(?:" + "|".join(map(re.escape, v)) + ")")
            for k, v in DATA_SOURCE.items()}
AUTO_MAP_DICT = {}
for (r_type, main_cat, proj), sub_list in RAW_MAPPING_CONFIG.items():
    for sub in sub_list:
        AUTO_MAP_DICT[sub] = (r_type, main_cat, proj)

COLUMN_MAPPING = {
    '交易时间': '交易时间', '交易创建时间': '交易时间', '付款时间': '交易时间',
    '交易类型': '交易类型', '交易分类': '交易类型', '交易对方': '交易对方',
    '商品': '商品', '商品名称': '商品', '收/支': '收/支',
    '金额(元)': '金额(元)', '金额（元）': '金额(元)', '金额': '金额(元)',
    '支付方式': '支付方式', '收/付款方式': '支付方式',
    '当前状态': '当前状态', '交易状态': '当前状态',
    '备注': '备注', '交易单号': '交易单号', '交易订单号': '交易单号',
    '商户单号': '商户单号', '商家订单号': '商户单号', '商品说明': '商品'
}

VALID_SUBCATS = sorted(list(AUTO_MAP_DICT.keys()), key=len, reverse=True)


# ==========================================
#          [核心] 逻辑处理函数
# ==========================================

def robust_date_converter(x):
    if isinstance(x, (datetime.datetime, datetime.date)):
        return pd.to_datetime(x)
    try:
        if isinstance(x, (int, float)):
            return pd.to_datetime(x, unit='D', origin='1899-12-30')
        return pd.to_datetime(x, errors='coerce')
    except Exception:
        return pd.NaT


def construct_description(df):
    df['描述'] = ""
    mask_tb = df['商家'] == '淘宝'
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
    if df.empty:
        return df
    return df


def load_settings(rule_path: Path):
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
    print("正在加载规则...", flush=True)
    try:
        with pd.ExcelFile(rule_path, engine='openpyxl') as xl:
            df = pd.read_excel(xl, sheet_name='Moze Dict')
        df['is_regex'] = pd.to_numeric(
            df.get('is_regex', 0), errors='coerce').fillna(0)
        df['商家(old)'] = df['商家(old)'].astype(str).str.strip()
        print(f"  已加载 {len(df)} 条规则")
        return df
    except Exception as e:
        print(f"  加载失败: {e}")
        return None


def sniff_and_load_data(file_path: Path):
    print(f"读取: {file_path.name}", flush=True)
    ftype, df = None, None
    try:
        if file_path.suffix.lower() == '.csv':
            df = pd.read_csv(file_path, header=24, encoding='gb18030')
            ftype = "AliPay"
        elif file_path.suffix.lower() == '.xlsx':
            df = pd.read_excel(file_path, header=16, engine='openpyxl')
            ftype = "WeChat"
        if df is None:
            return None

        df.columns = df.columns.astype(str).str.strip().str.replace(
            '[（(]', '(', regex=True).str.replace('[）)]', ')', regex=True)
        df.rename(columns=COLUMN_MAPPING, inplace=True)
        if '金额(元)' not in df.columns:
            return None
        df['金额(元)'] = pd.to_numeric(df['金额(元)'].astype(str).str.replace(
            r'[¥,]', '', regex=True).str.strip(), errors='coerce').fillna(0)
        df['金额'] = df['金额(元)']
        df['_source_tag'] = '#WechatPay' if ftype == "WeChat" else '#AliPay'
        print(f"  成功 (类型: {ftype})")
        return df
    except Exception as e:
        print(f"  读取失败: {e}")
        return None


def ensure_columns(df, main_col, sub_col):
    cols_to_check = [main_col, sub_col, '金额', '记录类型', '账户', '商家',
                     '描述', '项目', '名称', '对象', '标签', '日期', '时间', '币种', '手续费', '折扣']
    for c in cols_to_check:
        if c not in df.columns:
            df[c] = pd.NA
    cols_to_fill = ['记录类型', '项目', '对象', main_col, sub_col, '商家', '名称']
    df[cols_to_fill] = df[cols_to_fill].fillna("")
    return df


def process_transfers(df, main_col, sub_col):
    if df.empty:
        return pd.DataFrame()

    # 1. 获取配置变量
    t1 = CONFIG['TRANSFER_TARGET_1']
    t2 = CONFIG['TRANSFER_TARGET_2']
    t_sb = CONFIG['TRANSFER_TARGET_SNOWBALL']  # <--- [引用] 引用配置

    # 2. 生成掩码
    mask_t1 = df["交易对方"] == t1
    mask_t2 = df["交易对方"] == t2
    mask_sb = df["交易对方"] == t_sb  # <--- [筛选]

    mask_in = df["收/支"] == "收入"
    mask_out = df["收/支"] == "支出"

    res_list = []

    def create_records(mask, out_acc, in_acc, sub_val='转账'):
        if not mask.any():
            return
        base = df[mask].copy()

        # 转出
        o_rec = base.copy()
        o_rec['金额'] = base['金额'] * -1
        o_rec['记录类型'] = '转出'
        o_rec['账户'] = out_acc
        o_rec[sub_col] = sub_val

        # 转入
        i_rec = base.copy()
        i_rec['记录类型'] = '转入'
        i_rec['账户'] = in_acc
        i_rec[sub_col] = sub_val

        res_list.extend([o_rec, i_rec])

    # 3. 执行生成逻辑
    create_records(mask_t1 & mask_in, '零钱2', '零钱3')
    create_records(mask_t1 & mask_out, '零钱3', '零钱2')
    create_records(mask_t2, '零钱3', '工商银行', '提现')
    create_records(mask_sb & mask_out, '平安银行4946', '武汉通', '充值')

    if not res_list:
        return pd.DataFrame()

    ret = pd.concat(res_list, ignore_index=True)
    ret = ensure_columns(ret, main_col, sub_col)
    ret[main_col] = "转账"
    ret['日期'] = ret['交易时间'].dt.strftime('%Y/%m/%d')
    ret['时间'] = ret['交易时间'].dt.strftime('%H:%M:%S')
    ret['_Sort_Date'] = ret['交易时间']
    ret[['币种', '手续费', '折扣']] = ["CNY", 0, 0]
    return ret


def process_heuristics(df, main_col, sub_col):
    """
    逻辑优先级：
    1. DATA_SOURCE (关键词推导) - 基础层 (不填名称)
    2. 子类/通用词 (正餐/日用) - 推导分类，不填名称，切分描述
    3. 报销词 - 填名称，填对象
    """
    df.reset_index(drop=True, inplace=True)

    df = ensure_columns(df, main_col, sub_col)
    cols_to_str = [sub_col, '项目', '描述', '商家', '商品']
    df[cols_to_str] = df[cols_to_str].astype(str).apply(
        lambda x: x.str.strip().replace("nan", ""))

    search_series = (
        df.get('商家(old)', '').astype(str) + " " +
        df['商家'].astype(str) + " " +
        df['商品'].astype(str) + " " +
        df['描述'].astype(str)
    )

    # =========================================================================
    # Phase 1: 基础推导 (DATA_SOURCE + INGREDIENT_PRIORITY)
    # =========================================================================
    uncat = df[sub_col] == ""

    # MEAL 检测（用于排除正餐场景）
    mask_meal = search_series.str.contains(PATTERNS['MEAL'], regex=True)

    # [BUG FIX] 对于INGREDIENTS，需要先检测是否有精确匹配食材关键词
    # 例如"生水饺"应该优先匹配INGREDIENTS而非MEAL中的"水饺"
    mask_ingredients_exact = search_series.str.contains(
        PATTERNS['INGREDIENTS'], regex=True)

    main_filter = uncat | (df[main_col].isin(['购物', '居家', '饮食']))
    for item in INGREDIENT_PRIORITY:
        # 支持3元组 (key, name, sub_c) 和 4元组 (key, name, sub_c, obj)
        key, name, sub_c = item[0], item[1], item[2]
        obj = item[3] if len(item) > 3 else None

        # 跳过没有对应PATTERNS的key
        if key not in PATTERNS:
            continue

        pat = PATTERNS[key]
        # 对于INGREDIENTS类别，不排除meal匹配（让精确的食材关键词优先）
        if key == 'INGREDIENTS':
            mask = search_series.str.contains(pat, regex=True) & main_filter
        else:
            # 对于其他类别，排除meal匹配，但如果同时匹配了INGREDIENTS则不排除
            mask = search_series.str.contains(pat, regex=True) & (
                (~mask_meal) | mask_ingredients_exact) & main_filter
        if key in ['SEAFOOD', 'PORK', 'POULTRY', 'BEEF_MUTTON', 'VEGETABLE']:
            mask &= (~search_series.str.contains(
                PATTERNS['COOKED'], regex=True))
        if mask.any():
            df.loc[mask, [sub_col, '名称']] = [sub_c, name]
            if obj:
                df.loc[mask, '对象'] = obj

    # 1.5 停车/借贷/报销
    mask = search_series.str.contains(PATTERNS['Parking_fee'], regex=True)
    if mask.any():
        df.loc[mask, [sub_col, '名称', '对象']] = ['报账', '停车费', '天之逸']

    # 借贷关键词：借入xxx, 借出xxx, 代付xxx, 押金xxx, 报账xxx
    debt_keys = ['报账', '借出', '代付', '押金', '借入']
    debt_pat = rf"({'|'.join(debt_keys)})\s*(.*)"
    extracted = df['描述'].str.extract(debt_pat, expand=True)
    mask_found = extracted[0].notna()
    if mask_found.any():
        df.loc[mask_found, sub_col] = extracted[0]
        mask_obj = mask_found & (extracted[1].str.strip() != "")
        if mask_obj.any():
            df.loc[mask_obj, '对象'] = extracted[1].str.strip()
        df.loc[mask_found, ['项目', '描述']] = ""

    # 报销关键词（无点语法）：住宿费xxx, 材料费xxx 等
    reim_travel_keys = ["车船费", "住宿费", "住宿补贴", "交通补贴", "餐费补贴"]
    reim_expense_keys = [
        "材料费", "燃油费", "交通费", "过路费", "租赁费",
        "叉车费", "停车费", "印刷服务", "物流运输", "市内交通",
        "生活用品", "人工劳务费", "代付货款", "招待费", "汽车费用"
    ]
    all_reim_keys = reim_travel_keys + reim_expense_keys
    reim_pat = rf"^({'|'.join(all_reim_keys)})(.*)"
    reim_extracted = df['描述'].str.extract(reim_pat, expand=True)
    mask_reim = reim_extracted[0].notna()
    if mask_reim.any():
        df.loc[mask_reim, sub_col] = '报账'
        df.loc[mask_reim, '对象'] = '天之逸'
        df.loc[mask_reim, '名称'] = reim_extracted[0].loc[mask_reim].values
        df.loc[mask_reim, '描述'] = reim_extracted[1].loc[mask_reim].str.strip().values
        df.loc[mask_reim, '项目'] = ""
        # 设置标签
        mask_travel = reim_extracted[0].isin(reim_travel_keys) & mask_reim
        mask_expense = reim_extracted[0].isin(reim_expense_keys) & mask_reim
        if mask_travel.any():
            df.loc[mask_travel, '标签'] = '差旅报销'
        if mask_expense.any():
            df.loc[mask_expense, '标签'] = '费用报销'

    # 1.5.2 通用分类词（无点语法统一管理）
    # 所有子类别 + 名称关键词，格式：关键词xxx → 子类别设置，描述为xxx

    # 子类别关键词映射
    subcat_keywords = {
        # 直接映射到子类别
        '日用': '日常用品',
        '食材': '食材',
        '零食': '零食',
        '饮料水果': '饮料水果',
        '纯净水': '纯净水',
        '早餐': '早餐',
        '午餐': '午餐',
        '晚餐': '晚餐',
        '夜宵': '夜宵',
    }

    # 名称关键词映射 (名称 → 子类别)
    name_keywords = {
        '水果': ('饮料水果', '水果'),
        '饮料': ('饮料水果', '饮料'),
        '蔬菜': ('食材', '蔬菜'),
        '猪肉': ('食材', '猪肉'),
        '牛羊肉': ('食材', '牛羊肉'),
        '禽肉': ('食材', '禽肉'),
        '海鲜水产': ('食材', '海鲜水产'),
        '豆制品': ('食材', '豆制品'),
        '熟食': ('食材', '熟食'),
        '大米': ('食材', '大米'),
    }

    # 合并所有关键词（按长度降序，优先匹配长的）
    all_generic_keys = list(subcat_keywords.keys()) + \
        list(name_keywords.keys()) + ['正餐']
    all_generic_keys = sorted(all_generic_keys, key=len, reverse=True)

    generic_pat = rf"^({'|'.join(map(re.escape, all_generic_keys))})(.*)"
    generic_extracted = df['描述'].str.extract(generic_pat, expand=True)
    mask_generic = generic_extracted[0].notna() & (df[sub_col] == "")

    if mask_generic.any():
        matched_keys = generic_extracted[0].loc[mask_generic]
        tails = generic_extracted[1].loc[mask_generic].str.strip()

        for idx in mask_generic[mask_generic].index:
            key = matched_keys.loc[idx]
            tail = tails.loc[idx]

            if key == '正餐':
                # 正餐特殊处理：只切分描述，不设置子类别，让它走时间推导
                df.loc[idx, '描述'] = tail
                df.loc[idx, '名称'] = ""
            elif key in subcat_keywords:
                # 子类别关键词：设置子类别
                df.loc[idx, sub_col] = subcat_keywords[key]
                df.loc[idx, '描述'] = tail
                df.loc[idx, '名称'] = ""
            elif key in name_keywords:
                # 名称关键词：设置子类别和名称
                sub_c, name = name_keywords[key]
                df.loc[idx, sub_col] = sub_c
                df.loc[idx, '名称'] = name
                df.loc[idx, '描述'] = tail

    # 1.6 MEAL 自动推导的时间段分类 (名称留空)
    # 特定商家（如幼儿园）由字典配置，这里统一做时间推导
    # 同时支持备注中的"正餐xxx"标记
    mask_meal_from_memo = df.get('_is_meal_from_memo', False) == True
    mask_time_meal = (df[sub_col] == "") & (mask_meal | mask_meal_from_memo)
    if mask_time_meal.any():
        h = df.loc[mask_time_meal, '交易时间'].dt.hour
        conditions = [(h >= 6) & (h < 11), (h >= 11) &
                      (h < 16), (h >= 16) & (h < 21)]
        choices = ["早餐", "午餐", "晚餐"]
        df.loc[mask_time_meal, sub_col] = np.select(
            conditions, choices, default="夜宵")

    mapped_values = df[sub_col].map(AUTO_MAP_DICT)
    mask_mapped = mapped_values.notna()
    if mask_mapped.any():
        df.loc[mask_mapped, '记录类型'] = mapped_values[mask_mapped].apply(
            lambda x: x[0])
        df.loc[mask_mapped, main_col] = mapped_values[mask_mapped].apply(
            lambda x: x[1])
        df.loc[mask_mapped, '项目'] = mapped_values[mask_mapped].apply(
            lambda x: x[2])

    df['描述'] = df['描述'].str.strip()
    return df


def process_main(df, df_rules, main_col, sub_col):
    for c in ['当前状态', '收/支', '交易对方', '交易时间', '备注']:
        if c not in df.columns:
            df[c] = ""
    df = df[
        (~df["当前状态"].isin(["已全额退款", "交易关闭"])) &
        (abs(df["金额"]) > 0.0001)
    ].copy()

    # [BUG FIX] 借入/借出等债务关键词检测：在收支筛选前先标记
    # 将备注中包含债务关键词的记录也保留下来，不只是"支出"
    debt_keywords = ['报账', '借出', '代付', '押金', '借入']
    debt_pattern = '|'.join(debt_keywords)
    memo_series = df['备注'].astype(str).str.strip()
    mask_has_debt_keyword = memo_series.str.contains(
        debt_pattern, na=False, regex=True)

    # 修改筛选条件：支出 OR 包含债务关键词
    df = df[(df["收/支"] == "支出") | mask_has_debt_keyword].copy()

    if df.empty:
        return pd.DataFrame()
    print(f"--- 处理主交易 ({len(df)}条) ---", flush=True)

    df = ensure_columns(df, main_col, sub_col)
    df['支付方式'] = df['支付方式'].astype(str).replace(
        STANDARDIZE_ACCOUNTS, regex=True)
    df['商家(old)'] = df.get('交易对方', pd.NA).astype(str).str.strip()
    rename_map = {c: f'{c}_rule' for c in df_rules.columns if c not in [
        '商家(old)', 'is_regex']}

    df = pd.merge(df, df_rules[df_rules['is_regex'] == 0].rename(columns=rename_map).drop(
        columns='is_regex', errors='ignore'), on='商家(old)', how='left')

    df.reset_index(drop=True, inplace=True)

    rule_col_check = f"{main_col}_rule"
    uncat_mask = df[rule_col_check].isna(
    ) if rule_col_check in df.columns else pd.Series(True, index=df.index)
    if uncat_mask.any():
        regex_rules = df_rules[df_rules['is_regex'] == 1]
        to_match = df.loc[uncat_mask, '交易对方'].astype(str)
        for _, r in regex_rules.iterrows():
            match_indices = to_match[to_match.str.contains(
                r['商家(old)'], regex=True)].index
            if not match_indices.empty:
                for c, rc in rename_map.items():
                    df.loc[match_indices, rc] = r[c]
                to_match = to_match.drop(match_indices)

    for orig, rule in rename_map.items():
        if rule in df.columns:
            if orig != '描述':
                df[orig] = df[rule].combine_first(df[orig])
            if orig != '描述':
                df.drop(columns=[rule], inplace=True)

    # [淘宝订单识别] 基于商户单号 T200P 前缀（T200P4/T200P5等）
    if '商户单号' in df.columns:
        mask_tb = (df.get('_source_tag', '') == '#AliPay') & df['商户单号'].astype(
            str).str.strip().str.startswith('T200P', na=False)
        if mask_tb.any():
            df.loc[mask_tb, '商家'] = '淘宝'

    df = construct_description(df)
    if '描述_rule' in df.columns:
        df.drop(columns=['描述_rule'], inplace=True)
    df['记录类型'] = df.get('记录类型', pd.NA).replace("", pd.NA).fillna(df['收/支'])
    df['商家'] = df.get('商家', pd.NA).replace("", pd.NA).fillna(df['交易对方'])
    df[sub_col] = df[sub_col].astype(str).str.strip().replace("nan", "")

    # [备注子类别解析] 格式：子类别+描述，如 "探索claude" → 子类=探索, 描述=claude
    # 特殊处理：日用→日常用品，正餐→时间推导
    df['_is_meal_from_memo'] = False  # 标记是否来自备注的正餐
    if '备注' in df.columns:
        memo_series = df['备注'].astype(str).str.strip().replace('nan', '')
        
        # 特殊关键词映射（不在AUTO_MAP_DICT中但需要支持的）
        special_keywords = {
            '日用': ('支出', '购物', '日常用品'),  # 日用 → 日常用品
        }
        
        # 先处理特殊关键词
        for keyword, (r_type, main_cat, subcat) in special_keywords.items():
            pattern = rf'^{re.escape(keyword)}(.*)$'
            matches = memo_series.str.match(pattern, na=False)
            if matches.any():
                extracted = memo_series[matches].str.replace(pattern, r'\1', regex=True).str.strip()
                df.loc[matches, '记录类型'] = r_type
                df.loc[matches, main_col] = main_cat
                df.loc[matches, sub_col] = subcat
                df.loc[matches, '描述'] = extracted
                memo_series[matches] = ''
        
        # 正餐特殊处理：只设置描述和标记，让后续时间推导决定子类别
        pattern = rf'^正餐(.*)$'
        matches = memo_series.str.match(pattern, na=False)
        if matches.any():
            extracted = memo_series[matches].str.replace(pattern, r'\1', regex=True).str.strip()
            df.loc[matches, '描述'] = extracted
            df.loc[matches, '_is_meal_from_memo'] = True  # 标记需要时间推导
            memo_series[matches] = ''
        
        # 处理标准子类别
        # 应收/应付款项的子类别列表（后面的内容应该是对象而不是描述）
        receivable_payable_subcats = {'借出', '代付', '报账', '押金', '借入'}
        
        for subcat in VALID_SUBCATS:
            if subcat not in AUTO_MAP_DICT:
                continue
            # 跳过正餐（已单独处理）
            if subcat == '正餐':
                continue
            # 匹配以子类别开头的备注
            pattern = rf'^{re.escape(subcat)}(.*)$'
            matches = memo_series.str.match(pattern, na=False)
            if matches.any():
                r_type, main_cat, proj = AUTO_MAP_DICT[subcat]
                # 提取后续内容
                extracted = memo_series[matches].str.replace(pattern, r'\1', regex=True).str.strip()
                # 设置映射值
                df.loc[matches, '记录类型'] = r_type
                df.loc[matches, main_col] = main_cat
                df.loc[matches, sub_col] = subcat
                # 应收/应付款项：后续内容是对象，不是描述
                if subcat in receivable_payable_subcats:
                    df.loc[matches, '对象'] = extracted
                    df.loc[matches, '描述'] = ""
                else:
                    df.loc[matches, '描述'] = extracted
                # 清空已处理的备注，避免后续重复处理
                memo_series[matches] = ''

    # 执行核心混合逻辑
    df = process_heuristics(df, main_col, sub_col)
    
    # 清理标记列
    if '_is_meal_from_memo' in df.columns:
        df.drop(columns=['_is_meal_from_memo'], inplace=True)

    mask_neg = (df['记录类型'].isin(['支出', '应付款项', '应收款项'])) & (
        df.get('收/支') == '支出')
    df.loc[mask_neg, '金额'] *= -1
    mask_borrow = df[sub_col] == '借入'
    df.loc[mask_borrow, '金额'] = df.loc[mask_borrow, '金额'].abs()

    mask_fix_pay = (df['支付方式'] == "/") & (df['记录类型'].isin(["收入", "应付款项"]))
    df.loc[mask_fix_pay, '支付方式'] = "零钱3"

    df['日期'] = df['交易时间'].dt.strftime('%Y/%m/%d')
    df['时间'] = df['交易时间'].dt.strftime('%H:%M:%S')
    df['_Sort_Date'] = df['交易时间']
    df['账户'] = df.get('支付方式', "")
    df[['币种', '手续费', '折扣']] = ["CNY", 0, 0]

    mask_debt = df['记录类型'].isin(['应收款项', '应付款项'])
    df.loc[mask_debt, '商家'] = ""
    df.loc[mask_debt, '项目'] = ""
    df = clean_final_description(df)
    return df


def get_user_input():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    root.lift()
    root.focus_force()
    root.update()
    print("\n选择文件...")
    files = filedialog.askopenfilenames(
        parent=root, title="请选择 Moze 导出的文件", filetypes=[("Excel/CSV", "*.xlsx;*.csv")])
    root.attributes('-topmost', False)
    root.update()
    if not files:
        root.destroy()
        return None, None
    start = simpledialog.askstring("筛选", "起始日期 yymmdd (留空导入全部):", parent=root)
    st_date = pd.to_datetime(start, format="%y%m%d") if start else None
    root.destroy()
    return files, st_date


def save_result(df, cols):
    if '标签' in df.columns:
        m = ~df['记录类型'].isin(['转入', '转出'])
        df.loc[m, '标签'] = (df.loc[m, '标签'].fillna(
            "") + " " + df.loc[m, '_source_tag']).str.strip()
    type_priority = {'转出': 0, '转入': 1}
    df['_Type_Rank'] = df['记录类型'].map(type_priority).fillna(0)
    df.sort_values(['_Sort_Date', '_Type_Rank'],
                   ascending=[False, True], inplace=True)
    if not TARGET_DIR.exists():
        TARGET_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    path = TARGET_DIR / f"MOZE导入_{timestamp}.csv"
    df.to_csv(path, index=False, columns=cols, encoding='utf-8-sig')
    return path


def main():
    print(f"{BColors.BOLD}=== Moze 导入脚本 v11.66 (Taobao Fix) ==={BColors.ENDC}")
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
            return

        files, st_date = get_user_input()
        if not files:
            return
        dfs = []
        for f in files:
            d = sniff_and_load_data(Path(f))
            if d is not None:
                dfs.append(d)
        if not dfs:
            return

        df_raw = pd.concat(dfs, ignore_index=True)
        df_raw['交易时间'] = df_raw['交易时间'].apply(robust_date_converter)
        if st_date:
            df_raw = df_raw[df_raw['交易时间'] >= st_date]

        # [核心] 更新筛选逻辑：包含雪球 target
        df_raw['交易对方'] = df_raw.get('交易对方', pd.NA).astype(str).str.strip()
        target_list = [
            CONFIG['TRANSFER_TARGET_1'],
            CONFIG['TRANSFER_TARGET_2'],
            CONFIG['TRANSFER_TARGET_SNOWBALL']  # <--- 雪
        ]
        mask_trans = df_raw['交易对方'].isin(target_list)

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
        for c in cols:
            if c not in df_final.columns:
                df_final[c] = ""
        path = save_result(df_final, cols)
        print(f"\n{BColors.OKGREEN}成功! 文件: {path}{BColors.ENDC}")

        checks = [
            ('支出/收入/转账', df_final['记录类型'].isin(['支出',
             '收入', '转入', '转出']), [main_col, sub_col]),
            ('应收/应付', df_final['记录类型'].isin(['应收款项', '应付款项']),
             [main_col, sub_col, '对象'])
        ]
        for name, mask, check_cols in checks:
            bad = mask & df_final[check_cols].isin(["", pd.NA]).any(axis=1)
            if bad.any():
                print(
                    f"{BColors.FAIL}[严重] {bad.sum()} 条【{name}】记录缺失关键信息:{BColors.ENDC}")
                display_cols = ['日期', '商家', '金额', '描述'] + check_cols
                print(df_final.loc[bad, display_cols].fillna(
                    '[空]').head(5).to_string(index=True))
        print(f"{BColors.OKGREEN}√ 处理完成{BColors.ENDC}")

    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    start_time = time.time()
    main()
    print(f"\n运行结束 | 总耗时: {time.time() - start_time:.2f} 秒")
    try:
        for i in range(3, 0, -1):
            print(f"\r程序将在 {i} 秒后自动关闭...", end="", flush=True)
            time.sleep(1)
    except KeyboardInterrupt:
        pass
