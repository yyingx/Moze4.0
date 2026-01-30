# -*- coding: utf-8 -*-
"""
Moze 导入脚本 v11.68 (Experimental - Full Refactor)
Created on Sun Jan 05 2026
Optimized: Thu Jan 30 2026

@author: TZY_YX

CHANGELOG (v11.68 Experimental):
[重构] 合并 DATA_SOURCE 和 INGREDIENT_PRIORITY 为统一的 INGREDIENT_CONFIG
[重构] 使用 logging 模块替代 print 语句
[重构] 提取所有魔法数字为常量
[修复] 动态检测 CSV/XLSX header 行数，适应格式变化
[修复] 改进异常处理，捕获具体异常类型
[修复] 添加日期/金额转换失败验证和日志
[修复] 点号分隔边界情况处理（空对象、空描述）
[优化] 向量化点号分隔处理，提升性能
[优化] 正则规则批量匹配优化
[新增] 添加 tqdm 进度条（可选）
[新增] 版本常量 __version__、__author__、__updated__
[移除] 删除 pd.options.mode.chained_assignment = None（危险设置）

基于 v11.67 的改进:
[保留] 点号分隔格式 → 备注"借入谢辉.钉子"自动解析为 对象=谢辉, 描述=钉子
[保留] 自动推断收/支 → 手动添加行时可不填"收/支"列
[保留] 报销关键词识别 → 材料费/燃油费/物流运输等
"""

# === 版本信息 ===
__version__ = '11.68-experimental'
__author__ = 'TZY_YX'
__updated__ = '2026-01-30'

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
import json

# === 日志配置 ===
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    datefmt='%H:%M:%S',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# 尝试导入 tqdm（可选）
try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False
    def tqdm(iterable, *args, **kwargs):
        return iterable


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
DATA_SOURCE_PATH = CURRENT_DIR / "data_source.json"  # 可选的外部配置

# --- 文件读取配置 ---
ALIPAY_HEADER_RANGE = (20, 30)  # 支付宝 header 搜索范围
WECHAT_HEADER_RANGE = (14, 20)  # 微信 header 搜索范围
ALIPAY_ENCODING = 'gb18030'
HEADER_KEYWORDS = ['交易时间', '付款时间', '交易创建时间']  # 用于检测 header

# --- 时间段配置 ---
MEAL_TIME_RANGES = {
    '早餐': (6, 11),
    '午餐': (11, 16),
    '晚餐': (16, 21),
    # 夜宵是默认值（21:00-6:00）
}

# --- 核心配置 ---
CONFIG = {
    # 转账对方识别
    'TRANSFER_TARGET_1': '肖恩',
    'TRANSFER_TARGET_2': '工商银行(9579)',
    'TRANSFER_TARGET_SNOWBALL': '上海雪球数智科技有限公司',
    # 转账账户映射
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

# --- 提取的常量 ---
DEBT_KEYWORDS = ['报账', '借出', '代付', '押金', '借入']
REIM_TRAVEL_KEYS = ["车船费", "住宿费", "住宿补贴", "交通补贴", "餐费补贴"]
REIM_EXPENSE_KEYS = [
    "材料费", "燃油费", "交通费", "过路费", "租赁费",
    "叉车费", "停车费", "印刷服务", "物流运输", "市内交通",
    "生活用品", "人工劳务费", "代付货款", "招待费", "汽车费用"
]
RECEIVABLE_PAYABLE_SUBCATS = {'借出', '代付', '报账', '押金', '借入'}

# 子类别关键词映射
SUBCAT_KEYWORDS = {
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

# 名称关键词映射 (名称 → (子类别, 名称))
NAME_KEYWORDS = {
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

# ==========================================
#    [重构] 统一的 INGREDIENT_CONFIG 数据结构
# ==========================================
# 合并原有的 DATA_SOURCE 和 INGREDIENT_PRIORITY
# 格式: key -> {keywords, name, subcategory, priority, object(可选)}

INGREDIENT_CONFIG = {
    # === 饮料水果 ===
    'DRINK': {
        'keywords': [
            "饮料", "可乐", "红牛", "奶茶", "东鹏", "果汁", "椰汁", "酸奶",
            "咖啡", "拿铁", "乐虎", "AD钙奶", "蜜雪冰城"
        ],
        'name': '饮料',
        'subcategory': '饮料水果',
        'priority': 1
    },
    'FRUIT': {
        'keywords': [
            "水果", "果园", "果蔬", "百果园", "鲜果", "果业", "苹果", "香蕉",
            "红枣", "大枣", "桃子", "水蜜桃", "樱桃", "黄桃", "香梨", "雪梨",
            "柑橘", "沃柑", "草莓", "甘蔗", "桔子", "沙糖桔", "葡萄", "西瓜",
            "柚子", "柠檬", "橙子", "菠萝", "榴莲", "山竹", "蓝莓", "荔枝",
            "龙眼", "哈密瓜", "椰子", "柿子", "杨梅", "李子", "芒果", "猕猴桃",
            "火龙果", "百香果", "莲雾", "车厘子", "菠萝蜜", "桑葚", "枇杷",
            "杨桃", "无花果", "圣女果", "小番茄", "姑娘果", "西梅", "青枣",
            "冬枣", "黑布林", "人参果", "丑八怪", "耙耙柑"
        ],
        'name': '水果',
        'subcategory': '饮料水果',
        'priority': 2
    },
    'WATER': {
        'keywords': ["纯净水", "矿泉水", "农夫山泉", "怡宝", "百岁山", "娃哈哈", "今麦郎"],
        'name': '',
        'subcategory': '纯净水',
        'priority': 3
    },

    # === 食材 ===
    'VEGETABLE': {
        'keywords': [
            "蔬菜", "白菜", "菠菜", "上海青", "油麦菜", "生菜", "娃娃菜", "空心菜",
            "苋菜", "菜芯", "韭菜", "香菜", "芹菜", "荠菜", "芥蓝", "土豆", "胡萝卜",
            "萝卜", "白萝卜", "青萝卜", "红薯", "紫薯", "山药", "芋头", "莲藕", "藕",
            "洋葱", "大蒜", "生姜", "竹笋", "芦笋", "茭白", "西红柿", "茄子", "黄瓜",
            "西葫芦", "南瓜", "冬瓜", "苦瓜", "丝瓜", "青椒", "彩椒", "尖椒", "秋葵",
            "玉米", "四季豆", "豇豆", "扁豆", "荷兰豆", "毛豆", "豌豆", "蚕豆", "刀豆",
            "香菇", "金针菇", "平菇", "杏鲍菇", "口蘑", "木耳", "银耳", "茶树菇",
            "猴头菇", "大葱", "小葱", "蒜苗", "蒜苔"
        ],
        'name': '蔬菜',
        'subcategory': '食材',
        'priority': 4
    },
    'RICE': {
        'keywords': ["大米", "五常"],
        'name': '大米',
        'subcategory': '食材',
        'priority': 5
    },
    'BEAN_PRODUCT': {
        'keywords': [
            "豆腐", "豆皮", "腐竹", "豆干", "豆腐泡", "千张", "豆卷", "腐皮",
            "油豆皮", "内脂豆腐", "老豆腐", "嫩豆腐", "冻豆腐", "响铃卷", "素鸡"
        ],
        'name': '豆制品',
        'subcategory': '食材',
        'priority': 6
    },
    'PORK': {
        'keywords': [
            "猪肉", "扇子骨", "板油", "五花肉", "梅花肉", "瘦肉", "猪蹄膀", "蹄膀",
            "排骨", "脊骨", "五花", "前蹄", "大肠", "筒子骨", "棒骨", "猪油", "猪蹄",
            "荤油", "鲜肉"
        ],
        'name': '猪肉',
        'subcategory': '食材',
        'priority': 7,
        'exclude_if_match': ['COOKED']  # 排除熟食
    },
    'BEEF_MUTTON': {
        'keywords': [
            "牛肉", "牛腩", "牛排", "牛柳", "牛肠", "牛杂", "牛腱", "肥牛",
            "羊肉", "羊排", "羊腿"
        ],
        'name': '牛羊肉',
        'subcategory': '食材',
        'priority': 8,
        'exclude_if_match': ['COOKED']
    },
    'POULTRY': {
        'keywords': [
            "三黄鸡", "鸡腿", "鸡翅", "鸡胸肉", "老母鸡", "乌鸡",
            "鸭肉", "鸭腿", "鸭翅", "鸭架", "老鸭"
        ],
        'name': '禽肉',
        'subcategory': '食材',
        'priority': 9,
        'exclude_if_match': ['COOKED']
    },
    'Eggs': {
        'keywords': ["鸡蛋", "土鸡蛋", "生咸鸭蛋", "皮蛋"],
        'name': '蛋及蛋制品',
        'subcategory': '食材',
        'priority': 10
    },
    'SEAFOOD': {
        'keywords': [
            "鱼", "虾", "蟹", "贝", "带鱼", "黄鱼", "鲈鱼", "鲫鱼", "草鱼",
            "基围虾", "皮皮虾", "大闸蟹", "生蚝", "鱿鱼", "章鱼", "海带", "紫菜"
        ],
        'name': '海鲜水产',
        'subcategory': '食材',
        'priority': 11,
        'exclude_if_match': ['COOKED']
    },
    'COOKED': {
        'keywords': [
            "熟食", "水煮花生", "花生米", "熟花生", "蚕豆", "毛豆", "藕夹", "茄盒",
            "锅包肉", "猪头肉", "卤菜", "凉菜", "烧鸡", "烤鸭", "酱牛肉", "熏鱼",
            "炸带鱼", "炸鱿鱼", "肉丸", "墨鱼丸", "牛肉丸", "鱼丸", "虾滑", "贡丸",
            "鱿鱼圈", "糍粑", "熟咸鸭蛋", "熟肉肠", "糯米制品", "铁板鸭"
        ],
        'name': '熟食',
        'subcategory': '食材',
        'priority': 12
    },
    'INGREDIENTS': {
        'keywords': [
            "食材", "馒头", "发糕", "生水饺", "鲜面条", "干面条", "挂面",
            "火腿", "腊肠", "榨菜", "甜酒", "老干妈", "杂酱", "酱料", "炸酱",
            "酱豆", "辣椒酱", "白砂糖", "食用盐", "生抽", "老抽", "耗油", "料酒",
            "胡椒粉", "辣椒粉", "蒸肉粉", "火锅底料", "红豆", "绿豆", "黄豆"
        ],
        'name': '',
        'subcategory': '食材',
        'priority': 13
    },

    # === 零食 ===
    'Snack': {
        'keywords': ["零食", "切糕", "蛋糕", "面包", "腰果"],
        'name': '',
        'subcategory': '零食',
        'priority': 14
    },

    # === 交通 ===
    'CHARGING': {
        'keywords': ["自助服务-充电桩"],
        'name': '充电',
        'subcategory': '加油充电',
        'priority': 15
    },

    # === 虚拟 ===
    'SOFTWARE': {
        'keywords': ["软件", "APP", "应用", "安卓"],
        'name': '',
        'subcategory': 'Software',
        'priority': 16
    },
    'SERVER': {
        'keywords': ["节点", "Dler", "Dogess"],
        'name': '节点',
        'subcategory': '虚拟其他',
        'priority': 17
    },

    # === 购物 ===
    'DAILY_NECESSITIES': {
        'keywords': [
            "日用", "抽纸", "卷纸", "厨房纸", "垃圾袋", "保鲜袋", "保鲜膜",
            "洗衣液", "洗洁精", "牙膏", "洗发水", "一次性手套", "一次性杯",
            "棉签", "纸巾"
        ],
        'name': '',
        'subcategory': '日常用品',
        'priority': 18
    },
    'Clothing_Shoes_Bags': {
        'keywords': ["袜子", "内裤", "帽子", "手套", "鞋", "T恤", "裤", "外套", "修裤脚"],
        'name': '',
        'subcategory': '服饰鞋包',
        'priority': 19
    },
    'Furniture_HomeTextiles': {
        'keywords': ["被子", "空调被", "枕头", "浴巾", "床笠"],
        'name': '',
        'subcategory': '家具家纺',
        'priority': 20
    },

    # === 医疗 ===
    'Adult_Products': {
        'keywords': ["避孕套", "成人润滑剂", "安全套", "Condoms"],
        'name': 'Condoms',
        'subcategory': '保健用品',
        'priority': 21
    },
}

# 正餐关键词（用于时间推导，不在 INGREDIENT_CONFIG 中）
MEAL_KEYWORDS = [
    "正餐", "东苑一层", "东苑二层", "西区食堂", "外勤", "斯迪姆幼儿园-柏思思",
    "三镇民生", "永和四喜", "老乡鸡", "黄蜀郎", "麦香园", "丝路",
    "兰州", "沙县", "长沙臭豆腐", "重庆小面", "麻辣香锅", "麻辣烫",
    "鸡公煲", "猪脚饭", "卤肉饭", "盖浇饭", "炒饭", "快餐", "小碗菜",
    "热干面", "太和板面", "板面", "油泼面", "牛肉面", "刀削面", "牛杂粉",
    "肠粉", "牛肉汤", "汤粉", "水煎包", "小笼包", "汽水包", "煎豆折",
    "鸡蛋饼", "包粑", "煎包", "煎饼", "烧饼", "锅盔", "肉夹馍", "馕",
    "水饺", "蒸饺", "混沌", "馄饨", "包子", "小面", "烧烤", "路边摊",
    "食堂", "餐厅", "早点", "小吃", "餐饮", "面馆"
]

# 特殊关键词
PARKING_KEYWORDS = ["WF7023"]

# ==========================================
#    [构建] 编译后的数据结构
# ==========================================

def build_data_structures():
    """从 INGREDIENT_CONFIG 构建运行时数据结构"""
    data_source = {}
    patterns = {}

    # 构建 DATA_SOURCE 和 PATTERNS
    for key, config in INGREDIENT_CONFIG.items():
        data_source[key] = config['keywords']
        patterns[key] = re.compile(r"(?:" + "|".join(map(re.escape, config['keywords'])) + ")")

    # 添加正餐和停车费
    data_source['MEAL'] = MEAL_KEYWORDS
    patterns['MEAL'] = re.compile(r"(?:" + "|".join(map(re.escape, MEAL_KEYWORDS)) + ")")
    data_source['Parking_fee'] = PARKING_KEYWORDS
    patterns['Parking_fee'] = re.compile(r"(?:" + "|".join(map(re.escape, PARKING_KEYWORDS)) + ")")

    # 添加报销关键词
    data_source['REIM_TRAVEL'] = REIM_TRAVEL_KEYS
    data_source['REIM_EXPENSE'] = REIM_EXPENSE_KEYS

    return data_source, patterns

DATA_SOURCE, PATTERNS = build_data_structures()

# 按优先级排序的食材配置
SORTED_INGREDIENT_CONFIG = sorted(
    INGREDIENT_CONFIG.items(),
    key=lambda x: x[1].get('priority', 999)
)

# RAW_MAPPING_CONFIG（保持不变）
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
ALL_REIM_KEYS = REIM_TRAVEL_KEYS + REIM_EXPENSE_KEYS
REIM_PATTERN = re.compile(rf"^({'|'.join(ALL_REIM_KEYS)})(.*)")
ALL_GENERIC_KEYS = sorted(
    list(SUBCAT_KEYWORDS.keys()) + list(NAME_KEYWORDS.keys()) + ['正餐'],
    key=len, reverse=True
)
GENERIC_PATTERN = re.compile(
    rf"^({'|'.join(map(re.escape, ALL_GENERIC_KEYS))})(.*)")


# ==========================================
#          [核心] 逻辑处理函数
# ==========================================

def robust_date_converter(x):
    """转换日期，增加日志记录"""
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
                df = pd.read_csv(file_path, header=i, encoding=encoding, nrows=1)
            else:
                df = pd.read_excel(file_path, header=i, engine='openpyxl', nrows=1)

            # 检查是否包含关键列名
            cols_str = ' '.join(df.columns.astype(str))
            if any(kw in cols_str for kw in HEADER_KEYWORDS):
                logger.debug(f"检测到 header 在第 {i} 行")
                return i
        except Exception:
            continue

    # 未找到，使用默认值
    logger.warning(f"未能自动检测 header，使用默认值: {start}")
    return start


def construct_description(df):
    """构建描述字段"""
    df = df.copy()
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
    """加载配置，改进异常处理"""
    if not rule_path.exists():
        logger.warning(f"配置文件不存在: {rule_path}")
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
    except ValueError as e:
        logger.debug(f"Settings sheet 不存在: {e}")
    except Exception as e:
        logger.warning(f"加载配置失败: {e}")


def load_rules(rule_path: Path):
    """加载规则，改进异常处理"""
    logger.info("正在加载规则...")

    if not rule_path.exists():
        logger.error(f"规则文件不存在: {rule_path}")
        logger.error(f"请确保文件位于: {rule_path.absolute()}")
        return None

    try:
        with pd.ExcelFile(rule_path, engine='openpyxl') as xl:
            if 'Moze Dict' not in xl.sheet_names:
                logger.error(f"规则文件中缺少 'Moze Dict' sheet")
                logger.error(f"可用的 sheet: {xl.sheet_names}")
                return None
            df = pd.read_excel(xl, sheet_name='Moze Dict')

        df['is_regex'] = pd.to_numeric(
            df.get('is_regex', 0), errors='coerce').fillna(0)
        df['商家(old)'] = df['商家(old)'].astype(str).str.strip()
        logger.info(f"已加载 {len(df)} 条规则")
        return df

    except PermissionError:
        logger.error(f"无法读取文件（可能被其他程序占用）: {rule_path}")
        return None
    except Exception as e:
        logger.error(f"加载规则失败: {type(e).__name__} - {e}")
        return None


def sniff_and_load_data(file_path: Path):
    """读取数据文件，动态检测 header"""
    logger.info(f"读取: {file_path.name}")
    ftype, df = None, None

    try:
        if file_path.suffix.lower() == '.csv':
            # 动态检测 header
            header = find_header_row(file_path, ALIPAY_HEADER_RANGE, ALIPAY_ENCODING)
            df = pd.read_csv(file_path, header=header, encoding=ALIPAY_ENCODING)
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

        # 金额转换并验证
        original_amounts = df['金额(元)'].copy()
        df['金额(元)'] = pd.to_numeric(
            df['金额(元)'].astype(str).str.replace(r'[¥,]', '', regex=True).str.strip(),
            errors='coerce'
        ).fillna(0)

        # 检测转换失败（原值非空且非0，但结果为0）
        failed_mask = (
            original_amounts.notna() &
            (original_amounts.astype(str).str.strip() != '') &
            (original_amounts.astype(str).str.strip() != '0') &
            (df['金额(元)'] == 0)
        )
        if failed_mask.any():
            logger.warning(f"{failed_mask.sum()} 条记录金额转换失败")

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
    df = df.copy()
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

    # 获取配置变量
    t1 = CONFIG['TRANSFER_TARGET_1']
    t2 = CONFIG['TRANSFER_TARGET_2']
    t_sb = CONFIG['TRANSFER_TARGET_SNOWBALL']
    acc_lq2 = CONFIG['ACCOUNT_LINGQIAN_2']
    acc_lq3 = CONFIG['ACCOUNT_LINGQIAN_3']
    acc_icbc = CONFIG['ACCOUNT_ICBC']
    acc_pingan = CONFIG['ACCOUNT_PINGAN']
    acc_wht = CONFIG['ACCOUNT_WUHANTONG']

    # 生成掩码
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

    # 执行生成逻辑
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
    """处理借贷关键词：借入xxx, 借出xxx, 代付xxx, 押金xxx, 报账xxx"""
    df = df.copy()
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
    """处理报销关键词：住宿费xxx, 材料费xxx 等"""
    df = df.copy()
    reim_extracted = df['描述'].str.extract(REIM_PATTERN, expand=True)
    mask_reim = reim_extracted[0].notna()
    if mask_reim.any():
        df.loc[mask_reim, sub_col] = '报账'
        df.loc[mask_reim, '对象'] = '天之逸'
        df.loc[mask_reim, '名称'] = reim_extracted[0].loc[mask_reim].values
        df.loc[mask_reim, '描述'] = reim_extracted[1].loc[mask_reim].str.strip().values
        df.loc[mask_reim, '项目'] = ""
        # 设置标签
        mask_travel = reim_extracted[0].isin(REIM_TRAVEL_KEYS) & mask_reim
        mask_expense = reim_extracted[0].isin(REIM_EXPENSE_KEYS) & mask_reim
        if mask_travel.any():
            df.loc[mask_travel, '标签'] = '#差旅报销'
        if mask_expense.any():
            df.loc[mask_expense, '标签'] = '#费用报销'
    return df


def process_generic_keywords(df, sub_col):
    """处理通用分类词（向量化版本）"""
    df = df.copy()
    generic_extracted = df['描述'].str.extract(GENERIC_PATTERN, expand=True)
    mask_generic = generic_extracted[0].notna() & (df[sub_col] == "")

    if not mask_generic.any():
        return df

    matched_keys = generic_extracted.loc[mask_generic, 0]
    tails = generic_extracted.loc[mask_generic, 1].str.strip()

    # 正餐特殊处理
    mask_meal = matched_keys == '正餐'
    if mask_meal.any():
        meal_idx = mask_meal[mask_meal].index
        df.loc[meal_idx, '描述'] = tails.loc[meal_idx]
        df.loc[meal_idx, '名称'] = ""

    # 子类别关键词处理（向量化）
    for keyword, subcat in SUBCAT_KEYWORDS.items():
        mask_kw = matched_keys == keyword
        if mask_kw.any():
            kw_idx = mask_kw[mask_kw].index
            df.loc[kw_idx, sub_col] = subcat
            df.loc[kw_idx, '描述'] = tails.loc[kw_idx]
            df.loc[kw_idx, '名称'] = ""

    # 名称关键词处理（向量化）
    for keyword, (subcat, name) in NAME_KEYWORDS.items():
        mask_kw = matched_keys == keyword
        if mask_kw.any():
            kw_idx = mask_kw[mask_kw].index
            df.loc[kw_idx, sub_col] = subcat
            df.loc[kw_idx, '名称'] = name
            df.loc[kw_idx, '描述'] = tails.loc[kw_idx]

    return df


def process_heuristics(df_in, main_col, sub_col):
    """
    启发式分类推导
    使用统一的 INGREDIENT_CONFIG 替代原有的双数据结构
    """
    df = df_in.copy()
    df.reset_index(drop=True, inplace=True)

    df = ensure_columns(df, main_col, sub_col)
    cols_to_str = [sub_col, '项目', '描述', '商家', '商品']
    for col in cols_to_str:
        if col in df.columns:
            df[col] = df[col].fillna('').astype(str).str.strip()

    search_series = (
        df.get('商家(old)', '').astype(str) + " " +
        df['商家'].astype(str) + " " +
        df['商品'].astype(str) + " " +
        df['描述'].astype(str)
    )

    # Phase 1: 基础推导 (使用 INGREDIENT_CONFIG)
    uncat = df[sub_col] == ""
    mask_meal = search_series.str.contains(PATTERNS['MEAL'], regex=True)

    # 预计算 COOKED 匹配（用于排除）
    mask_cooked = search_series.str.contains(PATTERNS.get('COOKED', re.compile('')), regex=True)

    main_filter = uncat | (df[main_col].isin(['购物', '居家', '饮食']))

    # 使用排序后的配置（按优先级）
    for key, config in tqdm(SORTED_INGREDIENT_CONFIG, desc="分类推导", leave=False, disable=not HAS_TQDM):
        if key not in PATTERNS:
            continue

        keywords = config['keywords']
        name = config.get('name', '')
        sub_c = config['subcategory']
        obj = config.get('object')
        exclude_if_match = config.get('exclude_if_match', [])

        pat = PATTERNS[key]

        if key == 'INGREDIENTS':
            mask = search_series.str.contains(pat, regex=True) & main_filter
        else:
            mask = search_series.str.contains(pat, regex=True) & (
                (~mask_meal) | (key == 'COOKED')) & main_filter

        # 排除已匹配熟食的情况
        if 'COOKED' in exclude_if_match:
            mask &= (~mask_cooked)

        if mask.any():
            df.loc[mask, [sub_col, '名称']] = [sub_c, name]
            if obj:
                df.loc[mask, '对象'] = obj

    # 停车费
    mask = search_series.str.contains(PATTERNS['Parking_fee'], regex=True)
    if mask.any():
        df.loc[mask, [sub_col, '名称', '对象']] = ['报账', '停车费', '天之逸']

    # 借贷/报销/通用分类词
    df = process_debt_keywords(df, sub_col)
    df = process_reimbursement(df, sub_col)
    df = process_generic_keywords(df, sub_col)

    # MEAL 自动推导的时间段分类（使用配置常量）
    mask_meal_from_memo = df.get('_is_meal_from_memo', False) == True
    mask_time_meal = (df[sub_col] == "") & (mask_meal | mask_meal_from_memo)
    if mask_time_meal.any():
        h = df.loc[mask_time_meal, '交易时间'].dt.hour
        conditions = []
        choices = []
        for meal_name, (start_h, end_h) in MEAL_TIME_RANGES.items():
            conditions.append((h >= start_h) & (h < end_h))
            choices.append(meal_name)
        df.loc[mask_time_meal, sub_col] = np.select(conditions, choices, default="夜宵")

    # 映射记录类型/主类别/项目
    mapped_values = df[sub_col].map(AUTO_MAP_DICT)
    mask_mapped = mapped_values.notna()
    if mask_mapped.any():
        df.loc[mask_mapped, '记录类型'] = mapped_values[mask_mapped].apply(lambda x: x[0])
        df.loc[mask_mapped, main_col] = mapped_values[mask_mapped].apply(lambda x: x[1])
        df.loc[mask_mapped, '项目'] = mapped_values[mask_mapped].apply(lambda x: x[2])

    df['描述'] = df['描述'].str.strip()
    return df


def parse_memo_subcategory_vectorized(df, main_col, sub_col):
    """
    [优化] 向量化版本的备注子类别解析
    支持点号分隔格式，并处理边界情况
    """
    if '备注' not in df.columns:
        return df

    df = df.copy()
    df['_is_meal_from_memo'] = False
    memo_series = df['备注'].astype(str).str.strip().replace('nan', '')

    # 特殊关键词映射
    special_keywords = {'日用': ('支出', '购物', '日常用品')}
    for keyword, (r_type, main_cat, subcat) in special_keywords.items():
        pattern = rf'^{re.escape(keyword)}(.*)$'
        matches = memo_series.str.match(pattern, na=False)
        if matches.any():
            extracted = memo_series[matches].str.replace(pattern, r'\1', regex=True).str.strip()
            df.loc[matches, '记录类型'] = r_type
            df.loc[matches, main_col] = main_cat
            df.loc[matches, sub_col] = subcat
            df.loc[matches, '描述'] = extracted
            memo_series = memo_series.where(~matches, '')

    # 正餐特殊处理
    pattern = rf'^正餐(.*)$'
    matches = memo_series.str.match(pattern, na=False)
    if matches.any():
        extracted = memo_series[matches].str.replace(pattern, r'\1', regex=True).str.strip()
        df.loc[matches, '描述'] = extracted
        df.loc[matches, '_is_meal_from_memo'] = True
        memo_series = memo_series.where(~matches, '')

    # 处理标准子类别（向量化版本）
    for subcat in VALID_SUBCATS:
        if subcat not in AUTO_MAP_DICT or subcat == '正餐':
            continue

        pattern = rf'^{re.escape(subcat)}(.*)$'
        matches = memo_series.str.match(pattern, na=False)

        if not matches.any():
            continue

        r_type, main_cat, proj = AUTO_MAP_DICT[subcat]
        extracted = memo_series[matches].str.replace(pattern, r'\1', regex=True).str.strip()

        df.loc[matches, '记录类型'] = r_type
        df.loc[matches, main_col] = main_cat
        df.loc[matches, sub_col] = subcat

        # [优化] 向量化点号分隔处理
        if subcat in RECEIVABLE_PAYABLE_SUBCATS:
            # 找出包含点号的记录
            mask_dot = extracted.str.contains(r'\.', regex=True, na=False)
            matched_indices = matches[matches].index

            # 处理包含点号的记录
            dot_indices = matched_indices[mask_dot[matched_indices]]
            if len(dot_indices) > 0:
                dot_content = extracted.loc[dot_indices]
                split_df = dot_content.str.split('.', n=1, expand=True)

                # 提取对象和描述
                obj_values = split_df[0].str.strip()
                desc_values = split_df[1].str.strip() if 1 in split_df.columns else pd.Series('', index=dot_indices)

                # [修复] 处理边界情况：如果对象为空，使用整个内容
                empty_obj_mask = (obj_values == '') | obj_values.isna()
                if empty_obj_mask.any():
                    # 对象为空时，使用原始内容（去掉点号）
                    obj_values = obj_values.where(
                        ~empty_obj_mask,
                        dot_content.str.replace('.', '', regex=False).str.strip()
                    )
                    desc_values = desc_values.where(~empty_obj_mask, '')

                df.loc[dot_indices, '对象'] = obj_values.values
                df.loc[dot_indices, '描述'] = desc_values.values

            # 处理不含点号的记录
            no_dot_indices = matched_indices[~mask_dot[matched_indices]]
            if len(no_dot_indices) > 0:
                df.loc[no_dot_indices, '对象'] = extracted.loc[no_dot_indices]
                df.loc[no_dot_indices, '描述'] = ""
        else:
            df.loc[matches, '描述'] = extracted

        memo_series = memo_series.where(~matches, '')

    return df


def apply_rules(df, df_rules, main_col, sub_col):
    """应用字典规则匹配"""
    df = df.copy()
    df['商家(old)'] = df.get('交易对方', pd.NA).astype(str).str.strip()
    rename_map = {c: f'{c}_rule' for c in df_rules.columns if c not in ['商家(old)', 'is_regex']}

    # 精确匹配
    df = pd.merge(
        df,
        df_rules[df_rules['is_regex'] == 0].rename(columns=rename_map).drop(columns='is_regex', errors='ignore'),
        on='商家(old)', how='left'
    )
    df.reset_index(drop=True, inplace=True)

    # 正则匹配（优化版本）
    rule_col_check = f"{main_col}_rule"
    uncat_mask = df[rule_col_check].isna() if rule_col_check in df.columns else pd.Series(True, index=df.index)

    if uncat_mask.any():
        regex_rules = df_rules[df_rules['is_regex'] == 1]
        if not regex_rules.empty:
            to_match = df.loc[uncat_mask, '交易对方'].astype(str)

            for _, r in regex_rules.iterrows():
                if to_match.empty:
                    break
                try:
                    match_mask = to_match.str.contains(r['商家(old)'], regex=True, na=False)
                    match_indices = to_match[match_mask].index
                    if not match_indices.empty:
                        for c, rc in rename_map.items():
                            df.loc[match_indices, rc] = r[c]
                        to_match = to_match.drop(match_indices)
                except re.error as e:
                    logger.warning(f"正则表达式错误 '{r['商家(old)']}': {e}")

    # 合并规则列
    for orig, rule in rename_map.items():
        if rule in df.columns and orig != '描述':
            df[orig] = df[rule].combine_first(df[orig])
            df.drop(columns=[rule], inplace=True)

    return df


def finalize_records(df, main_col, sub_col):
    """最终处理：金额符号、日期时间、账户等"""
    df = df.copy()

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
    """主交易处理入口"""
    df = df_in.copy()
    for c in ['当前状态', '收/支', '交易对方', '交易时间', '备注']:
        if c not in df.columns:
            df[c] = ""

    # 筛选有效交易
    df = df[
        (~df["当前状态"].isin(["已全额退款", "交易关闭"])) &
        (abs(df["金额"]) > 0.0001)
    ].copy()

    # [v11.67] 自动推断空白的收/支列
    all_reim_keywords = REIM_TRAVEL_KEYS + REIM_EXPENSE_KEYS
    reim_pattern = '|'.join(all_reim_keywords)

    if '备注' in df.columns:
        memo = df['备注'].astype(str).str.strip()
        mask_empty_inout = df['收/支'].fillna('').astype(str).str.strip().isin(['', 'nan', 'NaN', 'None'])

        # 借入 → 收入
        mask_borrow_in = mask_empty_inout & memo.str.contains(r'^借入', na=False, regex=True)
        if mask_borrow_in.any():
            df.loc[mask_borrow_in, '收/支'] = '收入'

        # 借出/代付/报账/押金 → 支出
        mask_borrow_out = mask_empty_inout & memo.str.contains(
            r'^(?:借出|代付|报账|押金)', na=False, regex=True)
        if mask_borrow_out.any():
            df.loc[mask_borrow_out, '收/支'] = '支出'

        # 报销关键词开头 → 支出
        mask_reim = mask_empty_inout & memo.str.contains(
            rf'^(?:{reim_pattern})', na=False, regex=True)
        if mask_reim.any():
            df.loc[mask_reim, '收/支'] = '支出'

    # 借入/借出等债务关键词检测：支出 OR 包含特殊关键词
    memo_series = df['备注'].astype(str).str.strip()
    all_special_pattern = '|'.join(DEBT_KEYWORDS) + '|' + reim_pattern
    mask_has_special_keyword = memo_series.str.contains(
        rf'^(?:{all_special_pattern})', na=False, regex=True)
    df = df[(df["收/支"] == "支出") | mask_has_special_keyword].copy()

    if df.empty:
        return pd.DataFrame()

    logger.info(f"处理主交易 ({len(df)} 条)")

    df = ensure_columns(df, main_col, sub_col)
    df['支付方式'] = df['支付方式'].astype(str).replace(STANDARDIZE_ACCOUNTS, regex=True)

    # 应用规则
    df = apply_rules(df, df_rules, main_col, sub_col)

    # 淘宝订单识别
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

    # [v11.68] 使用向量化版本的备注解析
    df = parse_memo_subcategory_vectorized(df, main_col, sub_col)

    # 执行核心混合逻辑
    df = process_heuristics(df, main_col, sub_col)

    # 清理标记列
    if '_is_meal_from_memo' in df.columns:
        df.drop(columns=['_is_meal_from_memo'], inplace=True)

    # 日期验证
    if df['交易时间'].isna().any():
        failed_count = df['交易时间'].isna().sum()
        logger.warning(f"{failed_count} 条记录日期解析失败，已跳过")
        df = df[df['交易时间'].notna()]

    # 最终处理
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
        parent=root,
        title="请选择 Moze 导出的文件",
        filetypes=[("Excel/CSV", "*.xlsx;*.csv")]
    )
    root.attributes('-topmost', False)
    root.update()

    if not files:
        root.destroy()
        return None, None

    start = simpledialog.askstring(
        "筛选",
        "起始日期 yymmdd (留空导入全部):",
        parent=root
    )
    st_date = pd.to_datetime(start, format="%y%m%d") if start else None
    root.destroy()
    return files, st_date


def save_result(df, cols):
    """保存结果"""
    df = df.copy()

    if '标签' in df.columns:
        m = ~df['记录类型'].isin(['转入', '转出'])
        df.loc[m, '标签'] = (
            df.loc[m, '标签'].fillna("") + " " + df.loc[m, '_source_tag']
        ).str.strip()

    type_priority = {'转出': 0, '转入': 1}
    df['_Type_Rank'] = df['记录类型'].map(type_priority).fillna(0)
    df.sort_values(['_Sort_Date', '_Type_Rank'], ascending=[False, True], inplace=True)

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
    print(f"  Moze 导入脚本 v{__version__}")
    print(f"  作者: {__author__} | 更新: {__updated__}")
    print("=" * 50)
    print(f"{BColors.ENDC}")

    if not HAS_TQDM:
        logger.info("提示: 安装 tqdm 可显示进度条 (pip install tqdm)")

    try:
        load_settings(RULE_BOOK_PATH)
        df_rules = load_rules(RULE_BOOK_PATH)
        if df_rules is None:
            input("按 Enter 退出")
            return

        cols = [c for c in df_rules.columns if c not in ['商家(old)', 'is_regex']]
        main_col = next((c for c in cols if "主类" in c), None)
        sub_col = next((c for c in cols if "子类" in c), None)
        if not main_col:
            logger.error("未找到主类别列")
            return

        files, st_date = get_user_input()
        if not files:
            return

        # 读取文件（带进度条）
        dfs = []
        for f in tqdm(files, desc="读取文件", disable=not HAS_TQDM):
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
            res_dfs.append(process_transfers(df_raw[mask_trans].copy(), main_col, sub_col))
        if (~mask_trans).any():
            res_dfs.append(process_main(df_raw[~mask_trans].copy(), df_rules, main_col, sub_col))

        if not res_dfs:
            logger.warning("无结果")
            return

        df_final = pd.concat(res_dfs, ignore_index=True)
        for c in cols:
            if c not in df_final.columns:
                df_final[c] = ""

        path = save_result(df_final, cols)
        print(f"\n{BColors.OKGREEN}✅ 成功! 文件: {path}{BColors.ENDC}")

        # 数据验证
        checks = [
            ('支出/收入/转账', df_final['记录类型'].isin(['支出', '收入', '转入', '转出']), [main_col, sub_col]),
            ('应收/应付', df_final['记录类型'].isin(['应收款项', '应付款项']), [main_col, sub_col, '对象'])
        ]

        has_issues = False
        for name, mask, check_cols in checks:
            bad = mask & df_final[check_cols].isin(["", pd.NA]).any(axis=1)
            if bad.any():
                has_issues = True
                print(f"{BColors.FAIL}[!] {bad.sum()} 条【{name}】记录缺失关键信息:{BColors.ENDC}")
                display_cols = ['日期', '商家', '金额', '描述'] + check_cols
                print(df_final.loc[bad, display_cols].fillna('[空]').head(5).to_string(index=True))

        if not has_issues:
            print(f"{BColors.OKGREEN}✅ 数据验证通过，无缺失{BColors.ENDC}")

        print(f"{BColors.OKGREEN}✅ 处理完成{BColors.ENDC}")

    except KeyboardInterrupt:
        logger.info("用户取消操作")
    except Exception as e:
        logger.error(f"发生错误: {type(e).__name__} - {e}")
        traceback.print_exc()


if __name__ == "__main__":
    start_time = time.time()
    main()
    elapsed = time.time() - start_time
    print(f"\n运行结束 | 总耗时: {elapsed:.2f} 秒")

    try:
        input("\n按 Enter 键退出...")
    except KeyboardInterrupt:
        pass
