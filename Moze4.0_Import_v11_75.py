# -*- coding: utf-8 -*-
"""
Moze 导入脚本 v11.75 (Income Keep Fix)
Created on Sun Jan 05 2026
Enhanced: Sat Feb 01 2026
Refactored: Mon Mar 02 2026
Fixed: Tue Mar 10 2026

@author: TZY_YX

基于 v11.74，合并 Trae 版 bug fix：
[修复] 收入记录被错误过滤 → process_main 过滤行改为保留收入
[修复] 收入类备注关键词（薪资/福利补贴等）无法自动推断收/支 → 新增 INFER_INCOME_PATTERN
[保留] v11.74 所有重构优化（子函数拆分、模块级预编译常量等）
"""

# === 版本信息 ===
__version__ = '11.75'
__author__ = 'TZY_YX'
__updated__ = '2026-03-10'

import numpy as np
import pandas as pd
from pathlib import Path
from tkinter import filedialog, simpledialog
import time
import tkinter as tk
import traceback
import re
import datetime
import logging

# === 日志配置 ===
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    datefmt='%H:%M:%S',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)


class BColors:
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    CYAN = '\033[96m'


# ==========================================
#          [配置] 常量和配置
# ==========================================

# --- 路径设置 ---
CURRENT_DIR = Path(__file__).parent if '__file__' in locals() else Path.cwd()
RULE_BOOK_PATH = CURRENT_DIR / "Moze Dict.xlsx"
TARGET_DIR = CURRENT_DIR / "Moze4.0_Import"

# --- 文件读取配置 ---
ALIPAY_HEADER_RANGE = (20, 30)
WECHAT_HEADER_RANGE = (14, 20)
ALIPAY_ENCODING = 'gb18030'
HEADER_KEYWORDS = ['交易时间', '付款时间', '交易创建时间']

# --- 核心配置 ---
CONFIG = {
    'TRANSFER_TARGET_1': '肖恩',
    'TRANSFER_TARGET_2': '工商银行(9579)',
    'TRANSFER_TARGET_SNOWBALL': '上海雪球数智科技有限公司',
    'ACCOUNT_LINGQIAN_2': '零钱2',
    'ACCOUNT_LINGQIAN_3': '零钱3',
    'ACCOUNT_ICBC': '工商银行',
    'ACCOUNT_PINGAN': '平安银行4946',
    'ACCOUNT_WUHANTONG': '武汉通',
}

STANDARDIZE_ACCOUNTS = {
    r'.*4946.*': '平安银行4946',
    r'.*9579.*': '工商银行',
    r'.*3379.*': '招商银行Ⅱ',
    r'.*4826.*': '广发银行4826',
    r'.*零钱.*': '零钱3'
}

# --- 关键词常量 ---
DEBT_KEYWORDS = ['报账', '借出', '代付', '押金', '借入']
REIM_TRAVEL_KEYS = ["车船费", "住宿费", "住宿补贴", "交通补贴", "餐费补贴"]
REIM_EXPENSE_KEYS = [
    "材料费", "燃油费", "交通费", "过路费", "租赁费",
    "叉车费", "停车费", "印刷服务", "物流运输", "市内交通",
    "生活用品", "人工劳务费", "代付货款", "招待费", "汽车费用"
]
# [v11.74] 模块级合并，消除 process_main 中的重复定义
ALL_REIM_KEYS = REIM_TRAVEL_KEYS + REIM_EXPENSE_KEYS
RECEIVABLE_PAYABLE_SUBCATS = {'借出', '代付', '报账', '押金', '借入'}

# 子类别关键词映射
SUBCAT_KEYWORDS = {
    '食材': '食材', '零食': '零食', '饮料水果': '饮料水果', '纯净水': '纯净水',
    '早餐': '早餐', '午餐': '午餐', '晚餐': '晚餐', '夜宵': '夜宵', '正餐': '正餐',
    '日用': '日常用品', '服饰': '服饰鞋包', '服饰鞋包': '服饰鞋包',
    '数码': '数码电器', '数码电器': '数码电器', '家具': '家具家纺', '家具家纺': '家具家纺', '大件': '大件',
    '加油充电': '加油充电', '公共交通': '公共交通',
    '火车': '火车', '机票': '机票', '出租车': '出租车', '停车费': '停车费',
    '房租': '房租', '水费': '水费', '电费': '电费', '物业费': '物业费', '宽带费': '宽带费',
    '快递': '快递邮政', '快递邮政': '快递邮政', '理发': '理发', '电话费': '电话费',
    'App': 'App', '订阅': '订阅', 'Software': 'Software', '软件': 'Software', '影音': '影音',
    '电影': '电影', '聚会': '聚会', '旅游': '旅游度假', '游戏': '网游电玩',
    '药品': '药品', '门诊': '门诊', '体检': '体检', '保健用品': '保健用品',
    '图书': '图书', '教材': '教材', '探索': '探索', '文具': '文具',
    '保险': '保险', '礼金': '礼金红包', '红包': '礼金红包',
    '薪资': '薪资', '福利补贴': '福利补贴', '年终奖': '年终奖', '收红包': '收红包',
    '利息收入': '利息收入', '投资盈利': '投资盈利', '二手折旧': '二手折旧', '其他收入': '其他收入',
}

NAME_KEYWORDS = {
    '水果': ('饮料水果', '水果'),
    '饮料': ('饮料水果', '饮料'),
    '充电': ('加油充电', '充电'),
    '加油': ('加油充电', '加油'),
    '蔬菜': ('食材', '蔬菜'),
    '猪肉': ('食材', '猪肉'),
    '牛羊肉': ('食材', '牛羊肉'),
    '禽肉': ('食材', '禽肉'),
    '海鲜水产': ('食材', '海鲜水产'),
    '豆制品': ('食材', '豆制品'),
    '熟食': ('食材', '熟食'),
    '大米': ('食材', '大米'),
    '鸡蛋': ('食材', '蛋及蛋制品'),
    '节点': ('虚拟其他', '节点'),
    '厨房用品': ('日常用品', '厨房用品')
}

# --- DATA_SOURCE ---
DATA_SOURCE = {
    # === 饮料水果 ===
    'DRINK': [
        "饮料",
        "可乐", "红牛", "奶茶", "东鹏", "果汁", "椰汁", "酸奶", "咖啡", "拿铁", "乐虎", "AD钙奶",
        "蜜雪冰城", "果粒橙"
    ],
    'FRUIT': [
        "水果",
        "果园", "果蔬", "百果园", "鲜果", "果业", "苹果", "香蕉", "红枣", "大枣", "桃子", "水蜜桃",
        "樱桃", "黄桃", "香梨", "雪梨", "柑橘", "沃柑", "草莓", "甘蔗", "桔子", "沙糖桔", "葡萄",
        "西瓜", "柚子", "柠檬", "橙子", "菠萝", "榴莲", "山竹", "蓝莓", "荔枝", "龙眼", "哈密瓜",
        "椰子", "柿子", "杨梅", "李子", "芒果", "猕猴桃", "火龙果", "百香果", "莲雾", "车厘子",
        "菠萝蜜", "桑葚", "枇杷", "杨桃", "无花果", "圣女果", "小番茄", "姑娘果",
        "西梅", "青枣", "冬枣", "黑布林", "人参果", "丑八怪", "耙耙柑"
    ],
    'WATER': [
        "纯净水",
        "矿泉水", "农夫山泉", "怡宝", "百岁山", "娃哈哈", "今麦郎"
    ],

    # === 食材 ===
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
        # 葱姜蒜/调味菜/快捷
        "大葱", "小葱", "蒜苗", "蒜苔", "方便面"
    ],
    'RICE': ["大米", "五常"],
    'BEAN_PRODUCT': [
        "豆腐", "豆皮", "腐竹", "豆干", "豆腐泡", "千张", "豆卷", "腐皮", "油豆皮", "内脂豆腐",
        "老豆腐", "嫩豆腐", "冻豆腐", "响铃卷", "素鸡"
    ],
    'PORK': [
        "猪肉", "扇子骨", "板油", "五花肉", "梅花肉", "瘦肉", "猪蹄膀", "蹄膀", "排骨", "脊骨",
        "五花", "前蹄", "大肠", "筒子骨", "棒骨", "猪油", "猪蹄", "荤油", "鲜肉"
    ],
    'BEEF_MUTTON': ["牛肉", "牛腩", "牛排", "牛柳", "牛肠", "牛杂", "牛腱", "肥牛", "羊肉", "羊排", "羊腿"],
    'POULTRY': ["冷鲜鸡", "鸡腿", "鸡翅", "鸡胸肉", "老母鸡", "乌鸡", "鸭肉", "鸭腿", "鸭翅", "鸭架", "老鸭"],
    'Eggs': ["鸡蛋", "土鸡蛋", "生咸鸭蛋", "皮蛋"],
    'SEAFOOD': [
        "鱼", "虾", "蟹", "贝", "带鱼", "黄鱼", "鲈鱼", "鲫鱼", "草鱼", "基围虾", "皮皮虾",
        "大闸蟹", "生蚝", "鱿鱼", "章鱼", "海带", "紫菜"
    ],
    'COOKED': [
        "熟食", "水煮花生", "花生米", "熟花生", "蚕豆", "毛豆", "藕夹", "茄盒", "锅包肉", "猪头肉", "卤菜", "凉菜", "烧鸡",
        "烤鸭", "酱牛肉", "熏鱼", "炸带鱼", "炸鱿鱼", "肉丸", "墨鱼丸", "牛肉丸", "鱼丸", "虾滑",
        "贡丸", "鱿鱼圈", "糍粑", "熟咸鸭蛋", "熟肉肠", "糯米制品", "铁板鸭", "鸭货"
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

    # === 零食 ===
    'Snack': [
        "零食",
        "切糕", "蛋糕", "面包", "腰果"
    ],

    # === 交通 ===
    'CHARGING': ["自助服务-充电桩"],

    # === 虚拟 ===
    'SOFTWARE': ["软件", "APP", "应用", "安卓"],
    'SERVER': ["节点", "Dler", "Dogess"],

    # === 购物 ===
    'DAILY_NECESSITIES': [
        "日用",
        "抽纸", "卷纸", "厨房纸", "垃圾袋", "保鲜袋", "保鲜膜", "洗衣液", "洗洁精", "牙膏",
        "洗发水", "一次性手套", "一次性杯", "棉签", "纸巾"
    ],
    'Clothing_Shoes_Bags': ["袜子", "内裤", "帽子", "手套", "鞋", "T恤", "裤", "外套", "修裤脚"],
    'Furniture_HomeTextiles': ["被子", "空调被", "枕头", "浴巾", "床笠"],

    # === 医疗 ===
    'Adult_Products': ["避孕套", "成人润滑剂", "安全套", "Condoms"],

    # === 正餐（用于时间推导，不在 INGREDIENT_PRIORITY 中）===
    'MEAL': [
        "正餐",
        # 1. 地点/校区
        "东苑一层", "东苑二层", "西区食堂", "竹一", "斯迪姆幼儿园-柏思思",
        "李李",
        # 2. 连锁品牌
        "三镇民生", "永和四喜", "老乡鸡", "黄蜀郎", "麦香园", "丝路",
        # 3. 强特征的风味/地域
        "兰州", "沙县", "长沙臭豆腐", "重庆小面",
        # 4. 具体餐品 - 饭/面/粉/锅
        "麻辣香锅", "麻辣烫", "鸡公煲", "猪脚饭", "卤肉饭", "盖浇饭", "炒饭", "快餐", "小碗菜",
        "热干面", "板面", "油泼面", "牛肉面", "刀削面", "牛杂粉", "肠粉", "牛肉汤", "汤粉", "粉面", "凉面", "凉皮",
        # 5. 具体餐品 - 面点/早餐/小吃
        "水煎包", "小笼包", "汽水包", "煎豆折", "鸡蛋饼", "包粑",
        "煎包", "煎饼", "烧饼", "锅盔", "肉夹馍", "馕",
        "水饺", "蒸饺", "混沌", "馄饨", "包子", "小面",
        # 6. 通用场景/店名
        "烧烤", "路边摊", "食堂", "餐厅", "早点", "小吃", "餐饮", "面馆"
    ],

    # === 其他（不在 INGREDIENT_PRIORITY 中）===
    'Parking_fee': ["WF7023"],
    'REIM_TRAVEL': REIM_TRAVEL_KEYS,
    'REIM_EXPENSE': REIM_EXPENSE_KEYS
}

# --- 2. INGREDIENT_PRIORITY ---
# 顺序与 DATA_SOURCE 对应，用于推导名称和子类别
# 格式: (key, 名称, 子类别, [排除列表])
INGREDIENT_PRIORITY = [
    # 饮料水果
    ('DRINK', '饮料', '饮料水果'),
    ('FRUIT', '水果', '饮料水果'),
    ('WATER', '', '纯净水'),
    # 食材
    ('VEGETABLE', '蔬菜', '食材', ['COOKED']),
    ('RICE', '大米', '食材'),
    ('BEAN_PRODUCT', '豆制品', '食材'),
    ('PORK', '猪肉', '食材', ['COOKED']),
    ('BEEF_MUTTON', '牛羊肉', '食材', ['COOKED']),
    ('POULTRY', '禽肉', '食材', ['COOKED']),
    ('Eggs', '蛋及蛋制品', '食材'),
    ('SEAFOOD', '海鲜水产', '食材', ['COOKED']),
    ('COOKED', '熟食', '食材'),
    ('INGREDIENTS', '', '食材'),
    # 零食
    ('Snack', '', '零食'),
    # 交通
    ('CHARGING', '充电', '加油充电'),
    # 虚拟
    ('SOFTWARE', '', 'Software'),
    ('SERVER', '节点', '虚拟其他'),
    # 购物
    ('DAILY_NECESSITIES', '', '日常用品'),
    ('Clothing_Shoes_Bags', '', '服饰鞋包'),
    ('Furniture_HomeTextiles', '', '家具家纺'),
    # 医疗
    ('Adult_Products', 'Condoms', '保健用品'),
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

# 预编译的正则表达式
DEBT_PATTERN = re.compile(rf"({'|'.join(DEBT_KEYWORDS)})\s*(.*)")
REIM_PATTERN = re.compile(rf"^({'|'.join(ALL_REIM_KEYS)})(.*)")
ALL_GENERIC_KEYS = sorted(
    list(SUBCAT_KEYWORDS.keys()) + list(NAME_KEYWORDS.keys()) + ['正餐'],
    key=len, reverse=True
)
GENERIC_PATTERN = re.compile(
    rf"^({'|'.join(map(re.escape, ALL_GENERIC_KEYS))})(.*)")

# [v11.74] 模块级预计算：process_main 中用于筛选的完整特殊关键词模式
_INCOME_KEYWORDS_EXCLUDE = {'薪资', '福利补贴', '年终奖'}
_EXPENSE_SUBCAT_KEYS = [
    k for k in SUBCAT_KEYWORDS if k not in _INCOME_KEYWORDS_EXCLUDE]
_REIM_PATTERN_STR = '|'.join(ALL_REIM_KEYS)
_SUBCAT_PATTERN_STR = '|'.join(map(re.escape, _EXPENSE_SUBCAT_KEYS))
_NAME_PATTERN_STR = '|'.join(map(re.escape, NAME_KEYWORDS.keys()))
ALL_SPECIAL_PATTERN = re.compile(
    rf"^(?:{'|'.join(DEBT_KEYWORDS)}|{_REIM_PATTERN_STR}|{_SUBCAT_PATTERN_STR}|{_NAME_PATTERN_STR})"
)
# 用于推断收/支的宽松版（包含收入类关键词）
_ALL_SPECIAL_FOR_INFER_STR = (
    '|'.join(DEBT_KEYWORDS) + '|' + _REIM_PATTERN_STR + '|' +
    _SUBCAT_PATTERN_STR + '|' + _NAME_PATTERN_STR
)
INFER_EXPENSE_PATTERN = re.compile(rf"^(?:{_ALL_SPECIAL_FOR_INFER_STR})")

# 用于推断收/支的分项模式（与 v11.73 保持一致）
_INFER_EXPENSE_SUBCAT_KEYS = [k for k in SUBCAT_KEYWORDS if k not in {
    '薪资', '福利补贴', '年终奖', '收红包', '利息收入', '投资盈利', '二手折旧', '其他收入'}]
INFER_SUBCAT_PATTERN = re.compile(
    rf"^(?:{'|'.join(map(re.escape, _INFER_EXPENSE_SUBCAT_KEYS))})")
INFER_NAME_PATTERN = re.compile(
    rf"^(?:{'|'.join(map(re.escape, NAME_KEYWORDS.keys()))})")

# [v11.75] 收入类备注关键词推断模式
_INCOME_KEYWORDS = {'薪资', '福利补贴', '年终奖', '收红包', '利息收入', '投资盈利', '二手折旧', '其他收入'}
INFER_INCOME_PATTERN = re.compile(
    rf"^(?:{'|'.join(map(re.escape, _INCOME_KEYWORDS))})")


# ==========================================
#          [核心] 逻辑处理函数
# ==========================================

def robust_date_converter(x):
    """转换日期，支持多种格式"""
    if isinstance(x, (datetime.datetime, datetime.date)):
        return pd.to_datetime(x)
    try:
        if isinstance(x, (int, float)):
            return pd.to_datetime(x, unit='D', origin='1899-12-30')
        return pd.to_datetime(x, errors='coerce')
    except Exception as e:
        logger.debug(f"日期转换失败: {x}, 错误: {e}")
        return pd.NaT


def find_header_row(file_path: Path, search_range: tuple, encoding: str = 'utf-8'):
    """动态检测 header 行号"""
    start, end = search_range
    for i in range(start, end + 1):
        try:
            if file_path.suffix.lower() == '.csv':
                df = pd.read_csv(file_path, header=i,
                                 encoding=encoding, nrows=1)
            else:
                df = pd.read_excel(file_path, header=i,
                                   engine='openpyxl', nrows=1)
            cols_str = ' '.join(df.columns.astype(str))
            if any(kw in cols_str for kw in HEADER_KEYWORDS):
                logger.debug(f"检测到 header 在第 {i} 行")
                return i
        except Exception:
            continue
    logger.debug(f"未能自动检测 header，使用默认值: {start}")
    return start


def construct_description(df):
    """构建描述字段"""
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


def load_settings(rule_path: Path):
    """加载配置"""
    if not rule_path.exists():
        return
    try:
        df = pd.read_excel(rule_path, sheet_name='Settings', engine='openpyxl')
        loaded_count = 0
        for _, row in df.iterrows():
            k, v = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
            if k in CONFIG and v and v.lower() != 'nan':
                CONFIG[k] = v
                loaded_count += 1
        if loaded_count > 0:
            logger.info(f"已加载 {loaded_count} 项自定义配置")
    except ValueError:
        pass
    except Exception as e:
        logger.warning(f"加载配置失败: {e}")


def load_rules(rule_path: Path):
    """加载规则"""
    logger.info("正在加载规则...")
    if not rule_path.exists():
        logger.error(f"规则文件不存在: {rule_path}")
        return None
    try:
        with pd.ExcelFile(rule_path, engine='openpyxl') as xl:
            if 'Moze Dict' not in xl.sheet_names:
                logger.error("规则文件中缺少 'Moze Dict' sheet")
                return None
            df = pd.read_excel(xl, sheet_name='Moze Dict')
        df['is_regex'] = pd.to_numeric(
            df.get('is_regex', 0), errors='coerce').fillna(0)
        df['商家(old)'] = df['商家(old)'].astype(str).str.strip()
        logger.info(f"已加载 {len(df)} 条规则")
        return df
    except PermissionError:
        logger.error(f"无法读取文件（可能被其他程序占用）: {rule_path}")
    except Exception as e:
        logger.error(f"加载规则失败: {type(e).__name__} - {e}")
    return None


def sniff_and_load_data(file_path: Path):
    """读取数据文件，动态检测 header"""
    logger.info(f"读取: {file_path.name}")
    ftype, df = None, None
    try:
        if file_path.suffix.lower() == '.csv':
            header = find_header_row(
                file_path, ALIPAY_HEADER_RANGE, ALIPAY_ENCODING)
            df = pd.read_csv(file_path, header=header,
                             encoding=ALIPAY_ENCODING)
            ftype = "AliPay"
        elif file_path.suffix.lower() == '.xlsx':
            header = find_header_row(file_path, WECHAT_HEADER_RANGE)
            df = pd.read_excel(file_path, header=header, engine='openpyxl')
            ftype = "WeChat"

        if df is None:
            logger.warning(f"无法读取文件: {file_path}")
            return None

        df.columns = df.columns.astype(str).str.strip().str.replace(
            '[（(]', '(', regex=True).str.replace('[）)]', ')', regex=True)
        df.rename(columns=COLUMN_MAPPING, inplace=True)

        if '金额(元)' not in df.columns:
            logger.error(f"文件缺少金额列: {file_path}")
            return None

        df['金额(元)'] = pd.to_numeric(
            df['金额(元)'].astype(str).str.replace(
                r'[¥,]', '', regex=True).str.strip(),
            errors='coerce'
        ).fillna(0)
        df['金额'] = df['金额(元)']
        df['_source_tag'] = '#WechatPay' if ftype == "WeChat" else '#AliPay'
        logger.info(f"成功 (类型: {ftype}, {len(df)} 条记录)")
        return df

    except FileNotFoundError:
        logger.error(f"文件不存在: {file_path}")
    except UnicodeDecodeError:
        logger.error(f"文件编码错误，请确保是 GB18030 编码: {file_path}")
    except pd.errors.ParserError as e:
        logger.error(f"文件格式错误: {e}")
    except Exception as e:
        logger.error(f"读取失败: {type(e).__name__} - {e}")
    return None


def ensure_columns(df, main_col, sub_col):
    """确保必要的列存在"""
    cols_to_check = [main_col, sub_col, '金额', '记录类型', '账户', '商家',
                     '描述', '项目', '名称', '对象', '标签', '日期', '时间', '币种', '手续费', '折扣']
    for c in cols_to_check:
        if c not in df.columns:
            df[c] = pd.NA
    cols_to_fill = ['记录类型', '项目', '对象', main_col, sub_col, '商家', '名称']
    df[cols_to_fill] = df[cols_to_fill].fillna("")
    return df


def process_transfers(df, main_col, sub_col):
    """处理转账记录"""
    if df.empty:
        return pd.DataFrame()

    t1 = CONFIG['TRANSFER_TARGET_1']
    t2 = CONFIG['TRANSFER_TARGET_2']
    t_sb = CONFIG['TRANSFER_TARGET_SNOWBALL']
    acc_lq2 = CONFIG['ACCOUNT_LINGQIAN_2']
    acc_lq3 = CONFIG['ACCOUNT_LINGQIAN_3']
    acc_icbc = CONFIG['ACCOUNT_ICBC']
    acc_pingan = CONFIG['ACCOUNT_PINGAN']
    acc_wht = CONFIG['ACCOUNT_WUHANTONG']

    mask_t1 = df["交易对方"] == t1
    mask_t2 = df["交易对方"] == t2
    mask_sb = df["交易对方"] == t_sb
    mask_in = df["收/支"] == "收入"
    mask_out = df["收/支"] == "支出"

    res_list = []

    def create_records(mask, out_acc, in_acc, sub_val='转账'):
        if not mask.any():
            return
        base = df[mask].copy()
        o_rec = base.copy()
        o_rec['金额'] = base['金额'] * -1
        o_rec['记录类型'] = '转出'
        o_rec['账户'] = out_acc
        o_rec[sub_col] = sub_val
        i_rec = base.copy()
        i_rec['记录类型'] = '转入'
        i_rec['账户'] = in_acc
        i_rec[sub_col] = sub_val
        res_list.extend([o_rec, i_rec])

    create_records(mask_t1 & mask_in, acc_lq2, acc_lq3)
    create_records(mask_t1 & mask_out, acc_lq3, acc_lq2)
    create_records(mask_t2, acc_lq3, acc_icbc, '提现')
    create_records(mask_sb & mask_out, acc_pingan, acc_wht, '充值')

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


def process_debt_keywords(df, sub_col):
    """处理借贷关键词"""
    extracted = df['描述'].str.extract(DEBT_PATTERN, expand=True)
    mask_found = extracted[0].notna()
    if mask_found.any():
        df.loc[mask_found, sub_col] = extracted[0]
        mask_obj = mask_found & (extracted[1].str.strip() != "")
        if mask_obj.any():
            df.loc[mask_obj, '对象'] = extracted[1].str.strip()
        df.loc[mask_found, ['项目', '描述']] = ""
    return df


def process_reimbursement(df, sub_col):
    """处理报销关键词"""
    reim_extracted = df['描述'].str.extract(REIM_PATTERN, expand=True)
    mask_reim = reim_extracted[0].notna()
    if mask_reim.any():
        df.loc[mask_reim, sub_col] = '报账'
        df.loc[mask_reim, '对象'] = '天之逸'
        df.loc[mask_reim, '名称'] = reim_extracted[0].loc[mask_reim].values
        df.loc[mask_reim, '描述'] = reim_extracted[1].loc[mask_reim].str.strip().values
        df.loc[mask_reim, '项目'] = ""
        mask_travel = reim_extracted[0].isin(REIM_TRAVEL_KEYS) & mask_reim
        mask_expense = reim_extracted[0].isin(REIM_EXPENSE_KEYS) & mask_reim
        if mask_travel.any():
            df.loc[mask_travel, '标签'] = '#差旅报销'
        if mask_expense.any():
            df.loc[mask_expense, '标签'] = '#费用报销'
    return df


def process_generic_keywords(df, sub_col):
    """处理通用分类词（向量化版本）"""
    generic_extracted = df['描述'].str.extract(GENERIC_PATTERN, expand=True)

    mask_subcat = generic_extracted[0].notna() & (df[sub_col] == "")
    mask_name = generic_extracted[0].notna(
    ) & generic_extracted[0].isin(NAME_KEYWORDS.keys())

    if mask_subcat.any():
        matched_keys = generic_extracted.loc[mask_subcat, 0]
        tails = generic_extracted.loc[mask_subcat, 1].str.strip()

        mask_meal = matched_keys == '正餐'
        if mask_meal.any():
            meal_idx = mask_meal[mask_meal].index
            df.loc[meal_idx, '描述'] = tails.loc[meal_idx]
            df.loc[meal_idx, '名称'] = ""

        for keyword, subcat in SUBCAT_KEYWORDS.items():
            mask_kw = matched_keys == keyword
            if mask_kw.any():
                kw_idx = mask_kw[mask_kw].index
                df.loc[kw_idx, sub_col] = subcat
                df.loc[kw_idx, '描述'] = tails.loc[kw_idx]
                df.loc[kw_idx, '名称'] = ""

    if mask_name.any():
        matched_keys_name = generic_extracted.loc[mask_name, 0]
        tails_name = generic_extracted.loc[mask_name, 1].str.strip()

        for keyword, (subcat, name) in NAME_KEYWORDS.items():
            mask_kw = matched_keys_name == keyword
            if mask_kw.any():
                kw_idx = mask_kw[mask_kw].index
                df.loc[kw_idx, sub_col] = subcat
                df.loc[kw_idx, '名称'] = name
                df.loc[kw_idx, '描述'] = tails_name.loc[kw_idx]

    return df


# ==========================================
# [v11.74] process_heuristics 拆分为三个子函数
# ==========================================

def _apply_ingredient_patterns(df, sub_col, search_series, main_col,
                               mask_meal, mask_ingredients_exact):
    """子函数①：依据 INGREDIENT_PRIORITY 匹配食材/商品类别"""
    main_filter = (df[sub_col] == "") | (df[main_col].isin(['购物', '居家', '饮食']))

    for item in INGREDIENT_PRIORITY:
        key, name, sub_c = item[0], item[1], item[2]
        exclude_list = item[3] if len(item) > 3 else []
        obj = item[4] if len(item) > 4 else None

        if key not in PATTERNS:
            continue

        pat = PATTERNS[key]
        if key == 'INGREDIENTS':
            mask = search_series.str.contains(pat, regex=True) & main_filter
        else:
            mask = (search_series.str.contains(pat, regex=True)
                    & (~mask_meal | mask_ingredients_exact)
                    & main_filter)

        for exclude_key in exclude_list:
            if exclude_key in PATTERNS:
                mask &= ~search_series.str.contains(
                    PATTERNS[exclude_key], regex=True)

        if mask.any():
            df.loc[mask, [sub_col, '名称']] = [sub_c, name]
            if obj:
                df.loc[mask, '对象'] = obj

    # 停车费（特殊规则）
    mask = search_series.str.contains(PATTERNS['Parking_fee'], regex=True)
    if mask.any():
        df.loc[mask, [sub_col, '名称', '对象']] = ['报账', '停车费', '天之逸']

    return df


def _infer_meal_time(df, sub_col, mask_meal, mask_meal_from_memo):
    """子函数②：根据交易时间推导正餐子类别"""
    mask_time_meal = (df[sub_col] == "") & (mask_meal | mask_meal_from_memo)
    if mask_time_meal.any():
        h = df.loc[mask_time_meal, '交易时间'].dt.hour
        conditions = [(h >= 6) & (h < 11), (h >= 11) &
                      (h < 16), (h >= 16) & (h < 21)]
        choices = ["早餐", "午餐", "晚餐"]
        df.loc[mask_time_meal, sub_col] = np.select(
            conditions, choices, default="夜宵")
    return df


def _map_record_types(df, sub_col, main_col):
    """子函数③：通过 AUTO_MAP_DICT 映射记录类型/主类别/项目"""
    mapped_values = df[sub_col].map(AUTO_MAP_DICT)
    mask_mapped = mapped_values.notna()
    if mask_mapped.any():
        df.loc[mask_mapped, '记录类型'] = mapped_values[mask_mapped].apply(
            lambda x: x[0])
        df.loc[mask_mapped, main_col] = mapped_values[mask_mapped].apply(
            lambda x: x[1])
        df.loc[mask_mapped, '项目'] = mapped_values[mask_mapped].apply(
            lambda x: x[2])
    return df


def process_heuristics(df_in, main_col, sub_col):
    """
    启发式分类推导（v11.74 重构）
    原单函数拆分为三个职责单一的子函数：
      ① _apply_ingredient_patterns  - 食材/商品类别匹配
      ② _infer_meal_time            - 正餐时间段推导
      ③ _map_record_types           - 记录类型映射
    """
    df = df_in.copy()
    df.reset_index(drop=True, inplace=True)

    cols_to_str = [sub_col, '项目', '描述', '商家', '商品']
    df[cols_to_str] = df[cols_to_str].astype(str).apply(
        lambda x: x.str.strip().replace("nan", ""))

    search_series = (
        df.get('商家(old)', '').astype(str) + " " +
        df['商家'].astype(str) + " " +
        df['商品'].astype(str) + " " +
        df['描述'].astype(str)
    )

    mask_meal = search_series.str.contains(PATTERNS['MEAL'], regex=True)
    mask_ingredients_exact = search_series.str.contains(
        PATTERNS['INGREDIENTS'], regex=True)

    # ① 食材/商品类别匹配
    df = _apply_ingredient_patterns(
        df, sub_col, search_series, main_col, mask_meal, mask_ingredients_exact)

    # 关键词处理（顺序与 v11.73 保持一致）
    df = process_debt_keywords(df, sub_col)
    df = process_reimbursement(df, sub_col)
    df = process_generic_keywords(df, sub_col)

    # ② 正餐时间段推导
    mask_meal_from_memo = df.get('_is_meal_from_memo', False) == True
    df = _infer_meal_time(df, sub_col, mask_meal, mask_meal_from_memo)

    # ③ 记录类型映射
    df = _map_record_types(df, sub_col, main_col)

    df['描述'] = df['描述'].str.strip()
    return df


def parse_memo_subcategory(df, main_col, sub_col):
    """解析备注中的子类别，向量化点号分隔处理"""
    if '备注' not in df.columns:
        return df

    df['_is_meal_from_memo'] = False
    memo_series = df['备注'].astype(str).str.strip().replace('nan', '')

    # 特殊关键词映射
    special_keywords = {'日用': ('支出', '购物', '日常用品')}
    for keyword, (r_type, main_cat, subcat) in special_keywords.items():
        pattern = rf'^{re.escape(keyword)}(.*)$'
        matches = memo_series.str.match(pattern, na=False)
        if matches.any():
            extracted = memo_series[matches].str.replace(
                pattern, r'\1', regex=True).str.strip()
            df.loc[matches, '记录类型'] = r_type
            df.loc[matches, main_col] = main_cat
            df.loc[matches, sub_col] = subcat
            df.loc[matches, '描述'] = extracted
            memo_series = memo_series.where(~matches, '')

    # 正餐特殊处理
    matches = memo_series.str.match(r'^正餐(.*)$', na=False)
    if matches.any():
        extracted = memo_series[matches].str.replace(
            r'^正餐(.*)$', r'\1', regex=True).str.strip()
        df.loc[matches, '描述'] = extracted
        df.loc[matches, '_is_meal_from_memo'] = True
        memo_series = memo_series.where(~matches, '')

    # 处理标准子类别
    for subcat in VALID_SUBCATS:
        if subcat not in AUTO_MAP_DICT or subcat == '正餐':
            continue
        pattern = rf'^{re.escape(subcat)}(.*)$'
        matches = memo_series.str.match(pattern, na=False)
        if not matches.any():
            continue

        r_type, main_cat, proj = AUTO_MAP_DICT[subcat]
        extracted = memo_series[matches].str.replace(
            pattern, r'\1', regex=True).str.strip()
        df.loc[matches, '记录类型'] = r_type
        df.loc[matches, main_col] = main_cat
        df.loc[matches, sub_col] = subcat

        if subcat in RECEIVABLE_PAYABLE_SUBCATS:
            matched_indices = matches[matches].index
            mask_dot = extracted.str.contains(r'\.', regex=True, na=False)

            dot_indices = matched_indices[mask_dot[matched_indices]]
            if len(dot_indices) > 0:
                dot_content = extracted.loc[dot_indices]
                split_df = dot_content.str.split('.', n=1, expand=True)
                obj_values = split_df[0].str.strip()
                desc_values = split_df[1].str.strip(
                ) if 1 in split_df.columns else pd.Series('', index=dot_indices)
                empty_obj_mask = (obj_values == '') | obj_values.isna()
                if empty_obj_mask.any():
                    obj_values = obj_values.where(
                        ~empty_obj_mask,
                        dot_content.str.replace(
                            '.', '', regex=False).str.strip()
                    )
                    desc_values = desc_values.where(~empty_obj_mask, '')
                df.loc[dot_indices, '对象'] = obj_values.values
                df.loc[dot_indices, '描述'] = desc_values.values

            no_dot_indices = matched_indices[~mask_dot[matched_indices]]
            if len(no_dot_indices) > 0:
                df.loc[no_dot_indices, '对象'] = extracted.loc[no_dot_indices]
                df.loc[no_dot_indices, '描述'] = ""
        else:
            df.loc[matches, '描述'] = extracted

        memo_series = memo_series.where(~matches, '')

    return df


def apply_rules(df, df_rules, main_col, sub_col):
    """应用字典规则匹配，添加正则异常处理"""
    df['商家(old)'] = df.get('交易对方', pd.NA).astype(str).str.strip()
    rename_map = {c: f'{c}_rule' for c in df_rules.columns if c not in [
        '商家(old)', 'is_regex']}

    df = pd.merge(
        df,
        df_rules[df_rules['is_regex'] == 0].rename(
            columns=rename_map).drop(columns='is_regex', errors='ignore'),
        on='商家(old)', how='left'
    )
    df.reset_index(drop=True, inplace=True)

    rule_col_check = f"{main_col}_rule"
    uncat_mask = df[rule_col_check].isna(
    ) if rule_col_check in df.columns else pd.Series(True, index=df.index)

    if uncat_mask.any():
        regex_rules = df_rules[df_rules['is_regex'] == 1]
        if not regex_rules.empty:
            to_match = df.loc[uncat_mask, '交易对方'].astype(str)
            for _, r in regex_rules.iterrows():
                if to_match.empty:
                    break
                try:
                    match_mask = to_match.str.contains(
                        r['商家(old)'], regex=True, na=False)
                    match_indices = to_match[match_mask].index
                    if not match_indices.empty:
                        for c, rc in rename_map.items():
                            df.loc[match_indices, rc] = r[c]
                        to_match = to_match.drop(match_indices)
                except re.error as e:
                    logger.warning(f"正则表达式错误 '{r['商家(old)']}': {e}")

    for orig, rule in rename_map.items():
        if rule in df.columns and orig != '描述':
            df[orig] = df[rule].combine_first(df[orig])
            df.drop(columns=[rule], inplace=True)

    return df


def finalize_records(df, main_col, sub_col):
    """最终处理：金额符号、日期时间、账户等"""
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

    return df


def process_main(df_in, df_rules, main_col, sub_col):
    """
    主交易处理入口（v11.74 重构）
    - 移除重复的 all_reim_keywords / reim_pattern 定义（已提升至模块级）
    - ensure_columns 只在入口处调用一次
    - 使用模块级 ALL_SPECIAL_PATTERN / INFER_SUBCAT_PATTERN / INFER_NAME_PATTERN
    """
    df = df_in.copy()
    for c in ['当前状态', '收/支', '交易对方', '交易时间', '备注']:
        if c not in df.columns:
            df[c] = ""

    df = df[
        (~df["当前状态"].isin(["已全额退款", "交易关闭"])) &
        (abs(df["金额"]) > 0.0001)
    ].copy()

    # 自动推断空白的收/支列
    if '备注' in df.columns:
        memo = df['备注'].astype(str).str.strip()
        mask_empty_inout = df['收/支'].fillna('').astype(
            str).str.strip().isin(['', 'nan', 'NaN', 'None'])

        mask_borrow_in = mask_empty_inout & memo.str.contains(
            r'^借入', na=False, regex=True)
        if mask_borrow_in.any():
            df.loc[mask_borrow_in, '收/支'] = '收入'

        mask_borrow_out = mask_empty_inout & memo.str.contains(
            r'^(?:借出|代付|报账|押金)', na=False, regex=True)
        if mask_borrow_out.any():
            df.loc[mask_borrow_out, '收/支'] = '支出'

        mask_reim = mask_empty_inout & memo.str.contains(
            rf"^(?:{'|'.join(ALL_REIM_KEYS)})", na=False, regex=True)
        if mask_reim.any():
            df.loc[mask_reim, '收/支'] = '支出'

        # [v11.74] 使用模块级预编译模式
        mask_subcat = mask_empty_inout & memo.str.contains(
            INFER_SUBCAT_PATTERN, na=False)
        mask_name_kw = mask_empty_inout & memo.str.contains(
            INFER_NAME_PATTERN, na=False)
        if mask_subcat.any():
            df.loc[mask_subcat, '收/支'] = '支出'
        if mask_name_kw.any():
            df.loc[mask_name_kw, '收/支'] = '支出'

        # [v11.75] 收入类备注关键词自动推断为收入
        mask_income = mask_empty_inout & memo.str.contains(
            INFER_INCOME_PATTERN, na=False)
        if mask_income.any():
            df.loc[mask_income, '收/支'] = '收入'

    memo_series = df['备注'].astype(str).str.strip()
    # [v11.74] 使用模块级 ALL_SPECIAL_PATTERN，无需重建
    mask_has_special_keyword = memo_series.str.contains(
        ALL_SPECIAL_PATTERN, na=False)
    # [v11.75] 收入只保留备注命中收入关键词的记录，避免大量未分类收入混入
    mask_income_with_keyword = (df["收/支"] == "收入") & memo_series.str.contains(
        INFER_INCOME_PATTERN, na=False)
    df = df[(df["收/支"] == "支出") | mask_income_with_keyword |
            mask_has_special_keyword].copy()

    if df.empty:
        return pd.DataFrame()

    logger.info(f"处理主交易 ({len(df)} 条)")

    # [v11.74] ensure_columns 只在此处调用一次
    df = ensure_columns(df, main_col, sub_col)
    df['支付方式'] = df['支付方式'].astype(str).replace(
        STANDARDIZE_ACCOUNTS, regex=True)

    df = apply_rules(df, df_rules, main_col, sub_col)

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

    df = parse_memo_subcategory(df, main_col, sub_col)
    df = process_heuristics(df, main_col, sub_col)

    if '_is_meal_from_memo' in df.columns:
        df.drop(columns=['_is_meal_from_memo'], inplace=True)

    if df['交易时间'].isna().any():
        failed_count = df['交易时间'].isna().sum()
        logger.warning(f"{failed_count} 条记录日期解析失败，已跳过")
        df = df[df['交易时间'].notna()]

    df = finalize_records(df, main_col, sub_col)
    return df


def get_user_input():
    """获取用户输入"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    root.lift()
    root.focus_force()
    root.update()

    logger.info("请选择账单文件...")
    files = filedialog.askopenfilenames(
        parent=root, title="请选择账单文件", filetypes=[("Excel/CSV", "*.xlsx;*.csv")])
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
    """保存结果"""
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
    """主函数"""
    print(f"{BColors.BOLD}{BColors.CYAN}")
    print("=" * 50)
    print(f"  Moze 导入脚本 v{__version__} (Income Keep Fix)")
    print(f"  作者: {__author__} | 更新: {__updated__}")
    print("=" * 50)
    print(f"{BColors.ENDC}")

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
            logger.error("未找到主类别列")
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
            logger.error("没有成功读取任何文件")
            return

        df_raw = pd.concat(dfs, ignore_index=True)
        df_raw['交易时间'] = df_raw['交易时间'].apply(robust_date_converter)
        if st_date:
            df_raw = df_raw[df_raw['交易时间'] >= st_date]

        df_raw['交易对方'] = df_raw.get('交易对方', pd.NA).astype(str).str.strip()
        target_list = [
            CONFIG['TRANSFER_TARGET_1'],
            CONFIG['TRANSFER_TARGET_2'],
            CONFIG['TRANSFER_TARGET_SNOWBALL']
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
            logger.warning("无结果")
            return

        df_final = pd.concat(res_dfs, ignore_index=True)
        for c in cols:
            if c not in df_final.columns:
                df_final[c] = ""

        path = save_result(df_final, cols)
        print(f"\n{BColors.OKGREEN}✅ 成功! 文件: {path}{BColors.ENDC}")

        checks = [
            ('支出/收入/转账', df_final['记录类型'].isin(['支出',
             '收入', '转入', '转出']), [main_col, sub_col]),
            ('应收/应付', df_final['记录类型'].isin(['应收款项', '应付款项']),
             [main_col, sub_col, '对象'])
        ]

        has_issues = False
        for name, mask, check_cols in checks:
            bad = mask & df_final[check_cols].isin(["", pd.NA]).any(axis=1)
            if bad.any():
                has_issues = True
                print(
                    f"{BColors.FAIL}[!] {bad.sum()} 条【{name}】记录缺失关键信息:{BColors.ENDC}")
                display_cols = ['日期', '商家', '金额', '描述'] + check_cols
                print(df_final.loc[bad, display_cols].fillna(
                    '').head(5).to_string(index=True))

        if not has_issues:
            print(f"{BColors.OKGREEN}✅ 数据验证通过{BColors.ENDC}")

        print(f"{BColors.OKGREEN}✅ 处理完成{BColors.ENDC}")

    except KeyboardInterrupt:
        logger.info("用户取消操作")
    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    start_time = time.time()
    main()
    print(f"\n运行结束 | 总耗时: {time.time() - start_time:.2f} 秒")

    try:
        input("\n按 Enter 键退出...")
    except (KeyboardInterrupt, EOFError):
        pass
