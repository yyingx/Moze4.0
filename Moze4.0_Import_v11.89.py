# -*- coding: utf-8 -*-

# 版本信息
__version__ = '11.89'
__author__ = 'TZY_YX'
__updated__ = '2026-06-14'

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
import csv
import io

# 日志
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


# 配置

# 路径
CURRENT_DIR = Path(__file__).parent if '__file__' in locals() else Path.cwd()
RULE_BOOK_PATH = CURRENT_DIR / "Moze Dict.xlsx"
TARGET_DIR = CURRENT_DIR / "Moze4.0_Import"
PAUSE_ON_EXIT = False

# 文件读取
ALIPAY_HEADER_RANGE = (20, 30)
WECHAT_HEADER_RANGE = (14, 20)
ALIPAY_ENCODING = 'gb18030'
HEADER_KEYWORDS = ['交易时间', '付款时间', '交易创建时间']

# 默认账户和转账对象
CONFIG = {
    'TRANSFER_TARGET_1': '肖恩',
    'TRANSFER_TARGET_2': '工商银行(9579)',
    'TRANSFER_TARGET_SNOWBALL': '上海雪球数智科技有限公司',
    'ACCOUNT_LINGQIAN_2': '零钱2',
    'ACCOUNT_LINGQIAN_3': '零钱3',
    'ACCOUNT_ICBC': '工商银行',
    'ACCOUNT_PINGAN': '平安银行4946',
    'ACCOUNT_WUHANTONG': '武汉通',
    'ACCOUNT_ALIPAY_BALANCE': 'Alipay_余额',
}

STANDARDIZE_ACCOUNTS = {
    r'.*4946.*': '平安银行4946',
    r'.*9579.*': '工商银行',
    r'.*3379.*': '招商银行Ⅱ',
    r'.*4826.*': '广发银行4826',
    r'.*零钱.*': '零钱3'
}

# 手动备注关键词
DEBT_KEYWORDS = ['报账', '借出', '代付', '押金', '借入']
REIM_TRAVEL_KEYS = ["车船费", "住宿费", "住宿补贴", "交通补贴", "餐费补贴"]
REIM_EXPENSE_KEYS = [
    "材料费", "燃油费", "交通费", "过路费", "租赁费",
    "叉车费", "停车费", "印刷服务", "物流运输", "市内交通",
    "生活用品", "人工劳务费", "代付货款", "招待费", "汽车费用"
]

ALL_REIM_KEYS = REIM_TRAVEL_KEYS + REIM_EXPENSE_KEYS
RECEIVABLE_PAYABLE_SUBCATS = {'借出', '代付', '报账', '押金', '借入'}

# 子类别关键词映射
SUBCAT_KEYWORDS = {
    '食材': '食材', '零食': '零食', '饮料水果': '饮料水果', '纯净水': '纯净水',
    '早餐': '早餐', '午餐': '午餐', '晚餐': '晚餐', '夜宵': '夜宵',
    '日用': '日常用品', '服饰': '服饰鞋包', '服饰鞋包': '服饰鞋包',
    '数码': '数码电器', '数码电器': '数码电器', '家具': '家具家纺', '家具家纺': '家具家纺', '大件': '大件',
    '加油充电': '加油充电', '公共交通': '公共交通',
    '火车': '火车', '机票': '机票', '出租车': '出租车',
    '房租': '房租', '水费': '水费', '电费': '电费', '物业费': '物业费', '宽带费': '宽带费',
    '快递': '快递邮政', '快递邮政': '快递邮政', '理发': '理发', '电话费': '电话费',
    'App': 'App', '订阅': '订阅', 'Software': 'Software', '软件': 'Software', '影音': '影音',
    '电影': '电影', '聚会': '聚会', '旅游': '旅游度假', '游戏': '网游电玩',
    '药品': '药品', '门诊': '门诊', '体检': '体检', '保健用品': '保健用品',
    '图书': '图书', '教材': '教材', '探索': '探索', '文具': '文具',
    '保险': '保险', '礼金': '礼金红包', '红包': '礼金红包',
    '薪资': '薪资', '福利补贴': '福利补贴', '年终奖': '年终奖', '收红包': '收红包',
    '利息收入': '利息收入', '投资盈利': '投资盈利', '二手折旧': '二手折旧', '其他收入': '其他收入', '返利回馈': '返利回馈',
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

LOCATION_TAG_KEYWORDS = {
    '#外勤': ['湖北工业大学', '武汉大学']
}

# 自动分类关键词
DATA_SOURCE = {
    # 饮料水果
    'DRINK': [
        "饮料",
        "可乐", "红牛", "奶茶", "东鹏", "果汁", "椰汁", "酸奶", "咖啡", "拿铁", "乐虎", "AD钙奶",
        "蜜雪冰城", "果粒橙", "雪糕"
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

    # 食材
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
        "大葱", "小葱", "蒜苗", "蒜苔", "蒜坨"
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
        "熟食", "水煮花生", "花生米", "熟花生", "蚕豆", "熟毛豆", "藕夹", "茄盒", "锅包肉", "猪头肉", "卤菜", "凉菜", "烧鸡",
        "烤鸭", "酱牛肉", "熏鱼", "炸带鱼", "炸鱿鱼", "肉丸", "墨鱼丸", "牛肉丸", "鱼丸", "虾滑",
        "贡丸", "鱿鱼圈", "糍粑", "熟咸鸭蛋", "熟肉肠", "糯米制品", "铁板鸭", "鸭货"
    ],
    'INGREDIENTS': [
        # 通用
        "食材",
        # 面点主食
        "面粉", "馒头", "发糕", "生水饺", "鲜面条", "干面条", "挂面",
        # 腊味腌货
        "火腿", "腊肠", "榨菜", "甜酒", "酱菜",
        # 调味品和酱料
        "老干妈", "杂酱", "酱料", "炸酱", "酱豆", "辣椒酱",
        "白砂糖", "冰糖", "食盐", "生抽", "老抽", "耗油", "料酒",
        "胡椒粉", "辣椒粉", "蒸肉粉", "火锅底料",
        # 干货杂粮
        "红豆", "绿豆", "黄豆"
    ],

    # 零食
    'Snack': [
        "零食",
        "切糕", "蛋糕", "面包", "腰果"
    ],

    # 交通
    'CHARGING': ["自助服务-充电桩"],

    # 虚拟
    'SOFTWARE': ["软件", "APP", "应用", "安卓"],
    'SERVER': ["节点", "Dler", "Dogess"],

    # 购物
    'DAILY_NECESSITIES': [
        "日用",
        "抽纸", "卷纸", "厨房纸", "垃圾袋", "保鲜袋", "保鲜膜", "洗衣液", "洗洁精", "牙膏",
        "洗发水", "一次性手套", "一次性杯", "棉签", "纸巾"
    ],
    'Clothing_Shoes_Bags': ["袜子", "内裤", "帽子", "手套", "鞋", "T恤", "裤", "外套", "修裤脚"],
    'Furniture_HomeTextiles': ["被子", "空调被", "枕头", "浴巾", "床笠"],

    # 医疗
    'Adult_Products': ["避孕套", "成人润滑剂", "安全套", "Condoms"],

    # 正餐关键词，用于按时间推导早餐/午餐/晚餐/夜宵。
    'MEAL': [
        # 1. 地点/校区
        "东苑一层", "东苑二层", "西区食堂", "竹一", "斯迪姆幼儿园-柏思思",
        "李李",
        # 2. 连锁品牌
        "三镇民生", "永和四喜", "老乡鸡", "黄蜀郎", "麦香园", "丝路",
        # 3. 强特征的风味/地域
        "兰州", "沙县", "长沙臭豆腐", "重庆小面",
        # 4. 具体餐品 - 饭/面/粉/锅
        "麻辣香锅", "麻辣烫", "鸡公煲", "猪脚饭", "卤肉饭", "盖浇饭", "炒饭", "炒面", "快餐", "小碗菜",
        "热干面", "板面", "油泼面", "牛肉面", "刀削面", "牛杂粉", "肠粉", "牛肉汤", "汤粉", "粉面", "凉面", "凉皮",
        # 5. 具体餐品 - 面点/早餐/小吃
        "水煎包", "小笼包", "汽水包", "煎豆折", "鸡蛋饼", "包粑",
        "煎包", "煎饼", "烧饼", "锅盔", "肉夹馍", "馕",
        "水饺", "蒸饺", "混沌", "馄饨", "包子", "小面",
        # 6. 通用场景/店名
        "烧烤", "路边摊", "食堂", "餐厅", "早点", "小吃", "餐饮", "面馆", "餐馆"
    ],

    # 特殊关键词
    'Parking_fee': ["F7023"],
    'REIM_TRAVEL': REIM_TRAVEL_KEYS,
    'REIM_EXPENSE': REIM_EXPENSE_KEYS
}

# 分类匹配优先级：顺序会影响最终命中结果。
# 格式: (DATA_SOURCE key, 名称, 子类别, [排除 key])
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
    ('支出', '饮食', '食'): ['早餐', '午餐', '晚餐', '夜宵', '食材', '饮料水果', '纯净水', '零食'],
    ('支出', '购物', '日用&家用'): ['服饰鞋包', '日常用品', '大件', '共享租赁', '摄影文印', '数码电器', '家具家纺'],
    ('支出', '交通', '通信&交通'): ['加油充电', '公共交通', '共享交通', '火车', '出租车', '汽车', '轮渡', '机票', '交通违章', '维修保养'],
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
    ('收入', '收入', '兼职'): ['外卖跑腿(CNY)'],
    ('收入', '收入', ''): ['其他收入', '收红包', '二手折旧'],
    ('收入', '收入', '工作'): ['薪资', '福利补贴', '年终奖'],
    ('收入', '收入', '理财'): ['利息收入', '投资盈利'],
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

# 常用匹配模式
DEBT_PATTERN = re.compile(rf"^({'|'.join(DEBT_KEYWORDS)})\s*(.*)")
REIM_PATTERN = re.compile(rf"^({'|'.join(ALL_REIM_KEYS)})(.*)")
REIM_START_PATTERN = re.compile(
    rf"^(?:{'|'.join(map(re.escape, ALL_REIM_KEYS))})")
ALL_GENERIC_KEYS = sorted(
    list(SUBCAT_KEYWORDS.keys()) + list(NAME_KEYWORDS.keys()),
    key=len, reverse=True
)
GENERIC_PATTERN = re.compile(
    rf"^({'|'.join(map(re.escape, ALL_GENERIC_KEYS))})(.*)")

# 手动备注关键词入口，用于保留可导入记录。
_INCOME_KEYWORDS_EXCLUDE = {'薪资', '福利补贴', '年终奖'}
_EXPENSE_SUBCAT_KEYS = [
    k for k in SUBCAT_KEYWORDS if k not in _INCOME_KEYWORDS_EXCLUDE]
_REIM_PATTERN_STR = '|'.join(ALL_REIM_KEYS)
_SUBCAT_PATTERN_STR = '|'.join(map(re.escape, _EXPENSE_SUBCAT_KEYS))
_NAME_PATTERN_STR = '|'.join(map(re.escape, NAME_KEYWORDS.keys()))
ALL_SPECIAL_PATTERN = re.compile(
    rf"^(?:{'|'.join(DEBT_KEYWORDS)}|{_REIM_PATTERN_STR}|{_SUBCAT_PATTERN_STR}|{_NAME_PATTERN_STR})"
)
# 支出类备注关键词。
_INFER_EXPENSE_SUBCAT_KEYS = [k for k in SUBCAT_KEYWORDS if k not in {
    '薪资', '福利补贴', '年终奖', '收红包', '利息收入', '投资盈利', '二手折旧', '其他收入', '返利回馈'}]
INFER_SUBCAT_PATTERN = re.compile(
    rf"^(?:{'|'.join(map(re.escape, _INFER_EXPENSE_SUBCAT_KEYS))})")
INFER_NAME_PATTERN = re.compile(
    rf"^(?:{'|'.join(map(re.escape, NAME_KEYWORDS.keys()))})")

# 收入类备注关键词。
_INCOME_KEYWORDS = {'薪资', '福利补贴', '年终奖', '收红包',
                    '利息收入', '投资盈利', '二手折旧', '其他收入', '返利回馈'}
INFER_INCOME_PATTERN = re.compile(
    rf"^(?:{'|'.join(map(re.escape, _INCOME_KEYWORDS))})")


# 通用辅助常量
BLANK_STRINGS = {'', '/', 'nan', 'NaN', 'None', '<NA>', 'NaT'}
RAW_REQUIRED_COLUMNS = ['当前状态', '收/支', '交易对方', '交易时间', '备注']
OUTPUT_FILTER_COLUMNS = {'商家(old)', 'is_regex', '商品'}
INVALID_STATUSES = {'已全额退款', '交易关闭'}
FULL_REFUND_PATTERN = r'已\s*全额退款'
PARTIAL_REFUND_PATTERN = r'已\s*退款(?!.*全额)'
REFUND_ROW_PATTERN = r'退款成功|已\s*退款|^退款|退货退款'
RECORD_TYPES_NEED_DEFAULT_ACCOUNT = {'收入', '应付款项', '返利回馈'}


# 处理函数

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
    """
    构建描述字段。

    优先级：备注 > 淘宝商品名 > 字典描述。
    描述_rule 只在当前描述为空时补充，避免覆盖手动备注。
    """
    df['描述'] = ""

    mask_tb = df['商家'].fillna('').astype(str).eq('淘宝')
    if mask_tb.any():
        df.loc[mask_tb, '描述'] = df.loc[mask_tb, '商品'].fillna('').astype(str)

    if '备注' in df.columns:
        raw_memo = normalize_text_series(df['备注'])
        df['描述'] = df['描述'].where(raw_memo.eq(''), raw_memo)

    if '描述_rule' in df.columns:
        rule_desc = normalize_text_series(df['描述_rule'])
        current_desc = normalize_text_series(df['描述'])
        mask_fill_rule = current_desc.eq('') & rule_desc.ne('')
        df.loc[mask_fill_rule, '描述'] = rule_desc.loc[mask_fill_rule]

    df['描述'] = normalize_text_series(df['描述']).str.replace(
        r'[\r\n]+', ' ', regex=True
    ).str.strip()
    df.loc[df['描述'].isin(['/', 'nan']), '描述'] = ""
    return df


def clean_auto_descriptions(df):
    """清理由账单自动带入、但不适合作为 Moze 描述的文本。"""
    record_type = normalize_text_series(
        df.get('记录类型', pd.Series('', index=df.index)))
    counterparty = normalize_text_series(
        df.get('交易对方', pd.Series('', index=df.index)))
    desc = normalize_text_series(df.get('描述', pd.Series('', index=df.index)))

    mask_rebate_ad = (
        record_type.eq('返利回馈')
        & (
            desc.eq(counterparty)
            | desc.str.contains(r'现金红包|领更多|打开.*红包', regex=True, na=False)
        )
    )
    if mask_rebate_ad.any():
        df.loc[mask_rebate_ad, '描述'] = ""
    return df


def load_settings(rule_path: Path):
    """加载 Settings sheet 中的 CONFIG 覆盖项。"""
    if not rule_path.exists():
        return
    try:
        df = pd.read_excel(rule_path, sheet_name='Settings', engine='openpyxl')
        loaded_count = 0
        for _, row in df.iterrows():
            if len(row) < 2:
                continue
            k = str(row.iloc[0]).strip()
            v = str(row.iloc[1]).strip()
            if k in CONFIG and v and v.lower() != 'nan':
                CONFIG[k] = v
                loaded_count += 1
        if loaded_count > 0:
            logger.info(f"已加载 {loaded_count} 项自定义配置")
    except ValueError:
        # 没有 Settings sheet 时使用默认配置。
        pass
    except Exception as e:
        logger.warning(f"加载配置失败: {e}")


def load_rules(rule_path: Path):
    """加载 Moze Dict 规则表。"""
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

        if '商家(old)' not in df.columns:
            logger.error("规则文件缺少必需列: 商家(old)")
            return None

        if 'is_regex' not in df.columns:
            df['is_regex'] = 0
        df['is_regex'] = pd.to_numeric(
            df['is_regex'], errors='coerce').fillna(0).astype(int)
        df['商家(old)'] = normalize_text_series(df['商家(old)'])
        df = df[df['商家(old)'].ne('')].copy()

        logger.info(f"已加载 {len(df)} 条规则")
        return df
    except PermissionError:
        logger.error(f"无法读取文件（可能被其他程序占用）: {rule_path}")
    except Exception as e:
        logger.error(f"加载规则失败: {type(e).__name__} - {e}")
    return None


def sniff_and_load_data(file_path: Path):
    """读取账单文件，动态检测 header 并统一列名。"""
    logger.info(f"读取: {file_path.name}")
    ftype, df = None, None
    suffix = file_path.suffix.lower()
    try:
        if suffix == '.csv':
            header = find_header_row(
                file_path, ALIPAY_HEADER_RANGE, ALIPAY_ENCODING)
            df = pd.read_csv(file_path, header=header,
                             encoding=ALIPAY_ENCODING)
            ftype = "AliPay"
        elif suffix == '.xlsx':
            header = find_header_row(file_path, WECHAT_HEADER_RANGE)
            df = pd.read_excel(file_path, header=header, engine='openpyxl')
            ftype = "WeChat"
        else:
            logger.warning(f"不支持的文件类型: {file_path}")
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
        return pd.DataFrame(), pd.DataFrame()

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
    mask_type_cq = df["交易类型"] == "零钱充值"
    mask_type_tx = df["交易类型"] == "零钱提现"
    mask_in = df["收/支"] == "收入"
    mask_out = df["收/支"] == "支出"
    handled_mask = pd.Series(False, index=df.index)

    res_list = []

    def create_records(mask, out_acc, in_acc, sub_val='转账'):
        nonlocal handled_mask
        if not mask.any():
            return
        handled_mask |= mask
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
    create_records(mask_t2 & mask_type_tx, acc_lq3, acc_icbc, '提现')
    create_records(mask_t2 & mask_type_cq, acc_icbc, acc_lq3, '充值')
    create_records(mask_sb & mask_out, acc_pingan, acc_wht, '充值')

    if not res_list:
        return pd.DataFrame(), df.copy()

    ret = pd.concat(res_list, ignore_index=True)
    ret = ensure_columns(ret, main_col, sub_col)
    ret[main_col] = "转账"
    ret['日期'] = ret['交易时间'].dt.strftime('%Y/%m/%d')
    ret['时间'] = ret['交易时间'].dt.strftime('%H:%M:%S')
    ret['_Sort_Date'] = ret['交易时间']
    ret[['币种', '手续费', '折扣']] = ["CNY", 0, 0]
    unhandled = df[~handled_mask].copy()
    if not unhandled.empty:
        logger.warning(f"{len(unhandled)} 条疑似转账未命中细分规则，将按普通交易处理")
    return ret, unhandled


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


def process_location_tags(df, sub_col):
    """处理地点标签关键词"""
    exclude = set(DEBT_KEYWORDS + REIM_TRAVEL_KEYS + REIM_EXPENSE_KEYS)
    memo = df['备注'].astype(str).str.strip()
    mask_valid = ~df[sub_col].isin(exclude)
    for tag, locs in LOCATION_TAG_KEYWORDS.items():
        mask = memo.str.contains(
            '|'.join(map(re.escape, locs)), na=False) & mask_valid
        if mask.any():
            df.loc[mask, '标签'] = (
                df.loc[mask, '标签'].fillna('').str.strip() + ' ' + tag
            ).str.strip()
    return df


def process_generic_keywords(df, sub_col):
    """处理通用分类词（向量化版本）"""
    generic_extracted = df['描述'].str.extract(GENERIC_PATTERN, expand=True)

    mask_subcat = generic_extracted[0].notna() & (df[sub_col] == "")
    mask_name = (
        generic_extracted[0].notna()
        & (df[sub_col] == "")
        & generic_extracted[0].isin(NAME_KEYWORDS.keys())
    )

    if mask_subcat.any():
        matched_keys = generic_extracted.loc[mask_subcat, 0]
        tails = generic_extracted.loc[mask_subcat, 1].str.strip()

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


def _apply_ingredient_patterns(df, sub_col, search_series, main_col,
                               mask_meal, mask_ingredients_exact):
    """按关键词优先级匹配食材和商品类别。"""
    main_filter = df[sub_col] == ""

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

    # 停车费报账特殊规则。
    mask = search_series.str.contains(PATTERNS['Parking_fee'], regex=True)
    if mask.any():
        df.loc[mask, [sub_col, '名称', '对象']] = ['报账', '停车费', '天之逸']
        df.loc[mask, '标签'] = (
            df.loc[mask, '标签'].fillna('').str.strip() + ' #费用报销'
        ).str.strip()

    return df


def _infer_meal_time(df, sub_col, mask_meal):
    mask_time_meal = (df[sub_col] == "") & mask_meal
    if mask_time_meal.any():
        h = df.loc[mask_time_meal, '交易时间'].dt.hour
        conditions = [(h >= 6) & (h < 11), (h >= 11) &
                      (h < 16), (h >= 16) & (h < 21)]
        choices = ["早餐", "午餐", "晚餐"]
        df.loc[mask_time_meal, sub_col] = np.select(
            conditions, choices, default="夜宵")
    return df


def _map_record_types(df, sub_col, main_col):
    """按子类别映射记录类型、主类别和项目。"""
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
    """按商家、商品和备注推导分类信息。"""
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
    mask_dict_meal = (
        normalize_text_series(df[main_col]).eq('饮食')
        & is_blank_series(df[sub_col])
        & normalize_text_series(df['项目']).eq('食')
    )
    mask_ingredients_exact = search_series.str.contains(
        PATTERNS['INGREDIENTS'], regex=True)

    # 先处理手动备注关键词，手动备注优先于字典和启发式分类。
    df = process_debt_keywords(df, sub_col)
    df = process_reimbursement(df, sub_col)
    df = process_location_tags(df, sub_col)
    df = process_generic_keywords(df, sub_col)

    # 再用商品和商家关键词补空白分类。
    df = _apply_ingredient_patterns(
        df, sub_col, search_series, main_col,
        mask_meal | mask_dict_meal,
        mask_ingredients_exact
    )

    # 正餐按交易时间拆成早餐、午餐、晚餐或夜宵。
    mask_meal = mask_meal | (mask_dict_meal & is_blank_series(df[sub_col]))
    df = _infer_meal_time(df, sub_col, mask_meal)

    # 最后把子类别映射到 Moze 记录字段。
    df = _map_record_types(df, sub_col, main_col)

    df['描述'] = df['描述'].str.strip()
    return df


def parse_memo_subcategory(df, main_col, sub_col):
    """解析备注中的子类别，向量化点号分隔处理"""
    if '备注' not in df.columns:
        return df

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

    # 处理标准子类别
    for subcat in VALID_SUBCATS:
        if subcat not in AUTO_MAP_DICT or subcat in ALL_REIM_KEYS:
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
    """应用字典规则，支持商家简单规则与商家+商品复合规则。"""
    df['商家(old)'] = normalize_text_series(
        df.get('交易对方', pd.Series('', index=df.index)))
    if '商品' not in df.columns:
        df['商品'] = ''

    rename_map = {
        c: f'{c}_rule' for c in df_rules.columns if c not in OUTPUT_FILTER_COLUMNS}
    simple_rules, compound_rules = split_rules(df_rules)

    exact_rules = simple_rules[simple_rules['is_regex'] == 0]
    if not exact_rules.empty:
        df = pd.merge(
            df,
            exact_rules.rename(columns=rename_map).drop(
                columns=['is_regex', '商品'], errors='ignore'),
            on='商家(old)', how='left'
        )
        df.reset_index(drop=True, inplace=True)

    rule_col_check = f"{main_col}_rule"
    uncat_mask = df[rule_col_check].isna(
    ) if rule_col_check in df.columns else pd.Series(True, index=df.index)
    if uncat_mask.any():
        regex_rules = simple_rules[simple_rules['is_regex'] == 1]
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
                            val = r[c]
                            if pd.notna(val) and str(val).strip() not in ('', 'nan'):
                                df.loc[match_indices, rc] = val
                        rule_main_val = r.get(main_col)
                        if pd.notna(rule_main_val) and str(rule_main_val).strip() not in ('', 'nan'):
                            to_match = to_match.drop(match_indices)
                except re.error as e:
                    logger.warning(f"正则表达式错误 '{r['商家(old)']}': {e}")

    for orig, rule in rename_map.items():
        if rule in df.columns and orig != '描述':
            mask_fill = is_blank_series(df[orig]) & ~is_blank_series(df[rule])
            df.loc[mask_fill, orig] = df.loc[mask_fill, rule]
            df.drop(columns=[rule], inplace=True)

    if not compound_rules.empty:
        df = apply_compound_rules(df, compound_rules, rename_map)

    return df


def apply_pdd_inout_rule(df, main_col, sub_col):
    """拼多多收入按返利处理；拼多多支出保持支出。"""
    inout = normalize_text_series(df['收/支'])
    mask_pdd = (
        normalize_text_series(df['交易对方']).str.contains(
            '拼多多', regex=False, na=False)
        | normalize_text_series(df.get('商家', pd.Series('', index=df.index))).eq('拼多多')
    )
    rebate_cols = ['记录类型', main_col, sub_col]
    df.loc[mask_pdd & inout.eq('收入'), rebate_cols + ['商家', '项目']] = [
        '返利回馈', '返利回馈', '返利回馈', '拼多多', '返利']
    df.loc[
        mask_pdd & inout.eq('支出') & df[rebate_cols].eq('返利回馈').any(axis=1),
        rebate_cols + ['项目']
    ] = ''

    return df


def finalize_records(df, main_col, sub_col):
    """最终处理：金额符号、日期时间、账户等"""
    df['支付方式'] = normalize_text_series(df['支付方式']).replace(
        STANDARDIZE_ACCOUNTS, regex=True)

    mask_neg = (df['记录类型'].isin(['支出', '应付款项', '应收款项'])) & (
        df.get('收/支') == '支出')
    df.loc[mask_neg, '金额'] *= -1

    mask_borrow = df[sub_col] == '借入'
    df.loc[mask_borrow, '金额'] = df.loc[mask_borrow, '金额'].abs()

    df = fill_default_accounts_by_source(df)

    df['日期'] = df['交易时间'].dt.strftime('%Y/%m/%d')
    df['时间'] = df['交易时间'].dt.strftime('%H:%M:%S')
    df['_Sort_Date'] = df['交易时间']
    df['账户'] = df.get('支付方式', "")
    df[['币种', '手续费', '折扣']] = ["CNY", 0, 0]

    mask_debt = df['记录类型'].isin(['应收款项', '应付款项'])
    df.loc[mask_debt, '商家'] = ""
    df.loc[mask_debt, '项目'] = ""

    return df


def normalize_text_series(series):
    """统一把空值/nan 字符串清洗为空字符串。"""
    return series.fillna('').astype(str).str.strip().replace({
        'nan': '', 'NaN': '', 'None': '', '<NA>': '', '/': ''
    })


def is_blank_series(series):
    """判断 Series 中的空字符串、NA 和常见空值文本。"""
    return series.fillna('').astype(str).str.strip().isin(BLANK_STRINGS)


def fill_default_accounts_by_source(df):
    """按账单来源给无支付方式的收入类记录补默认账户。

    微信收入/应付款/返利常见支付方式为 `/`，支付宝同类记录常见为空；
    但这里不依赖空值形态，而是优先根据 `_source_tag` 判断来源。
    """
    need_default = (
        df['记录类型'].isin(RECORD_TYPES_NEED_DEFAULT_ACCOUNT)
        & is_blank_series(df['支付方式'])
    )
    if not need_default.any():
        return df

    source_tag = normalize_text_series(
        df.get('_source_tag', pd.Series('', index=df.index))
    )
    wechat_mask = need_default & (source_tag == '#WechatPay')
    alipay_mask = need_default & (source_tag == '#AliPay')
    unknown_mask = need_default & ~(wechat_mask | alipay_mask)

    if wechat_mask.any():
        df.loc[wechat_mask, '支付方式'] = CONFIG['ACCOUNT_LINGQIAN_3']
    if alipay_mask.any():
        df.loc[alipay_mask, '支付方式'] = CONFIG['ACCOUNT_ALIPAY_BALANCE']
    if unknown_mask.any():
        logger.warning(
            f"{unknown_mask.sum()} 条收入/应付款/返利记录缺少支付方式且来源未知，账户留空"
        )
    return df


def split_rules(df_rules):
    """把规则拆成简单规则与复合规则。"""
    if '商品' not in df_rules.columns:
        return df_rules.copy(), pd.DataFrame(columns=df_rules.columns)

    product = normalize_text_series(df_rules['商品'])
    compound_mask = product.ne('')
    return df_rules[~compound_mask].copy(), df_rules[compound_mask].copy()


def apply_compound_rules(df, compound_rules, rename_map):
    """应用商家+商品复合规则；描述写入描述_rule，商家不覆盖简单规则结果。"""
    bill_product = normalize_text_series(df['商品'])
    merchant_old = normalize_text_series(df['商家(old)'])

    for _, r in compound_rules.iterrows():
        merchant_pattern = str(r['商家(old)']).strip()
        product_pattern = str(r['商品']).strip()
        if (
            not merchant_pattern
            or not product_pattern
            or merchant_pattern.lower() == 'nan'
            or product_pattern.lower() == 'nan'
        ):
            continue

        use_regex = int(r.get('is_regex', 0)) == 1
        try:
            hit = (
                merchant_old.str.contains(merchant_pattern, regex=use_regex, na=False) &
                bill_product.str.contains(
                    product_pattern, regex=use_regex, na=False)
            )
        except re.error as e:
            logger.warning(
                f"复合规则正则错误 '{merchant_pattern}' / '{product_pattern}': {e}")
            continue

        if not hit.any():
            continue

        for c in rename_map:
            val = r.get(c)
            if pd.isna(val) or str(val).strip() in ('', 'nan'):
                continue

            if c == '描述':
                df.loc[hit, '描述_rule'] = val
            elif c == '商家':
                # 复合规则的商家只补空，避免覆盖简单规则清洗出的商家名。
                target_mask = hit & is_blank_series(df['商家'])
                df.loc[target_mask, '商家'] = val
            else:
                # 其他字段允许复合规则覆盖简单规则，体现更高优先级。
                df.loc[hit, c] = val
        logger.debug(
            f"复合规则 [{merchant_pattern} & 商品含'{product_pattern}']: {hit.sum()} 条")

    return df


def prepare_raw_transactions(df_in):
    """复制原始交易并补齐基础列。"""
    df = df_in.copy()
    for c in RAW_REQUIRED_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    return df


def warn_refund_rows(df_raw):
    """列出退款行，并标记需要手动处理的退货/部分退款。"""
    idx = df_raw.index
    inout = normalize_text_series(df_raw.get('收/支', pd.Series('', index=idx)))
    status = normalize_text_series(
        df_raw.get('当前状态', pd.Series('', index=idx)))
    trans_type = normalize_text_series(
        df_raw.get('交易类型', pd.Series('', index=idx)))
    memo = normalize_text_series(df_raw.get('备注', pd.Series('', index=idx)))
    product = normalize_text_series(df_raw.get('商品', pd.Series('', index=idx)))
    order_col = '交易单号' if '交易单号' in df_raw.columns else '交易订单号'
    order_no = normalize_text_series(
        df_raw.get(order_col, pd.Series('', index=idx)))

    refund_text = status + ' ' + trans_type + ' ' + product + ' ' + memo
    mask_refund = refund_text.str.contains(
        REFUND_ROW_PATTERN, regex=True, na=False)
    if not mask_refund.any():
        return

    base_order = order_no.str.extract(r'^(\d{8,})', expand=False).fillna('')
    refund_date = pd.to_datetime(df_raw.get('交易时间'), errors='coerce')
    order_date = pd.to_datetime(
        base_order.str.extract(r'^(\d{8})', expand=False),
        format='%Y%m%d',
        errors='coerce'
    )
    gap_days = (refund_date.dt.normalize() - order_date).dt.days

    mask_partial_income = inout.eq('收入') & refund_text.str.contains(
        PARTIAL_REFUND_PATTERN, regex=True, na=False
    )
    source_tag = normalize_text_series(
        df_raw.get('_source_tag', pd.Series('', index=idx))
    )
    mask_wechat = source_tag.eq('#WechatPay')
    mask_wechat_refund_expense = (
        mask_wechat
        & inout.eq('支出')
        & status.str.contains(PARTIAL_REFUND_PATTERN, regex=True, na=False)
    )
    mask_return_refund = mask_refund & (
        mask_partial_income | gap_days.isna() | (gap_days >= 2)
    )
    mask_return_refund &= ~mask_wechat_refund_expense
    manual_count = int(mask_return_refund.sum())

    if manual_count:
        print(f"{BColors.WARNING}⚠️ 请手动处理 {manual_count} 笔退款。{BColors.ENDC}")
    else:
        print(f"{BColors.WARNING}⚠️ 发现 {mask_refund.sum()} 笔退款，已按快速退款忽略。{BColors.ENDC}")

    display = pd.DataFrame({
        '来源': source_tag.replace({'#WechatPay': '微信', '#AliPay': '支付宝'}),
        '类型': np.select(
            [mask_return_refund, mask_wechat_refund_expense],
            [f"{BColors.FAIL}退货/部分退款{BColors.ENDC}", '原支出已部分退款'],
            default='快速退款'
        ),
        '交易时间': df_raw.get('交易时间'),
        '交易对方': df_raw.get('交易对方', ''),
        '金额': df_raw.get('金额', ''),
        '间隔天数': gap_days.astype('Int64').astype(str).replace('<NA>', ''),
    })
    rows = display.loc[mask_refund & ~
                       mask_wechat_refund_expense].fillna('').head(10)
    buffer = io.StringIO()
    rows.to_csv(buffer, index=False, lineterminator='\n',
                quoting=csv.QUOTE_MINIMAL)
    print(buffer.getvalue().rstrip())


def get_wechat_transfer_income_mask(df):
    """识别无备注的微信转账收入；这类记录只提示，不导入。"""
    idx = df.index
    source_tag = normalize_text_series(
        df.get('_source_tag', pd.Series('', index=idx))
    )
    inout = normalize_text_series(df.get('收/支', pd.Series('', index=idx)))
    trans_type = normalize_text_series(
        df.get('交易类型', pd.Series('', index=idx)))
    counterparty = normalize_text_series(
        df.get('交易对方', pd.Series('', index=idx))
    )
    memo = normalize_text_series(df.get('备注', pd.Series('', index=idx)))
    product = normalize_text_series(df.get('商品', pd.Series('', index=idx)))

    internal_targets = {
        CONFIG['TRANSFER_TARGET_1'],
        CONFIG['TRANSFER_TARGET_2'],
        CONFIG['TRANSFER_TARGET_SNOWBALL'],
    }
    return (
        source_tag.eq('#WechatPay')
        & inout.eq('收入')
        & trans_type.eq('转账')
        & memo.eq('')
        & (product + ' ' + memo).str.contains('微信转账', na=False)
        & ~counterparty.isin(internal_targets)
    )


def warn_wechat_transfer_income(df_raw):
    """提示无备注的微信转账收入，避免把语义不明的收款漏掉。"""
    idx = df_raw.index
    counterparty = normalize_text_series(
        df_raw.get('交易对方', pd.Series('', index=idx))
    )
    mask = get_wechat_transfer_income_mask(df_raw)
    if not mask.any():
        return

    print(f"{BColors.WARNING}⚠️ 请手动处理 {mask.sum()} 笔微信转账收入。{BColors.ENDC}")
    rows = pd.DataFrame({
        '来源': '微信',
        '类型': '转账收入',
        '交易时间': df_raw.get('交易时间'),
        '交易对方': counterparty,
        '金额': df_raw.get('金额', ''),
    }).loc[mask].head(10)
    buffer = io.StringIO()
    rows.to_csv(buffer, index=False, lineterminator='\n',
                quoting=csv.QUOTE_MINIMAL)
    print(buffer.getvalue().rstrip())


def infer_inout_from_memo(df):
    """根据备注关键词补齐空白的原始流水方向。

    这里的“收/支”只用于筛选记录和决定金额方向，不代表最终 Moze 记录类型。
    借入会先标为收入，后续再由子类别映射成“应付款项”。
    借出/代付/报账/押金会先标为支出，后续再映射成“应收款项”。
    """
    memo = normalize_text_series(df['备注'])
    mask_empty_inout = is_blank_series(df['收/支'])

    rules = [
        ('收入', mask_empty_inout & memo.str.contains(r'^借入', na=False, regex=True)),
        ('支出', mask_empty_inout & memo.str.contains(
            r'^(?:借出|代付|报账|押金)', na=False, regex=True)),
        ('支出', mask_empty_inout & memo.str.contains(REIM_START_PATTERN, na=False)),
        ('支出', mask_empty_inout & memo.str.contains(
            INFER_SUBCAT_PATTERN, na=False)),
        ('支出', mask_empty_inout & memo.str.contains(INFER_NAME_PATTERN, na=False)),
        ('收入', mask_empty_inout & memo.str.contains(
            INFER_INCOME_PATTERN, na=False)),
    ]
    for value, mask in rules:
        if mask.any():
            df.loc[mask, '收/支'] = value
    return df


def mark_taobao_orders(df):
    """支付宝淘宝订单统一标记商家。"""
    if '商户单号' not in df.columns:
        return df
    source_tag = df.get('_source_tag', pd.Series('', index=df.index))
    mask_tb = (source_tag == '#AliPay') & normalize_text_series(
        df['商户单号']).str.startswith('T200P', na=False)
    if mask_tb.any():
        df.loc[mask_tb, '商家'] = '淘宝'
    return df


def drop_invalid_dates(df):
    """删除日期解析失败的记录。"""
    if df['交易时间'].isna().any():
        failed_count = df['交易时间'].isna().sum()
        logger.warning(f"{failed_count} 条记录日期解析失败，已跳过")
        df = df[df['交易时间'].notna()].copy()
    return df


def get_output_columns(df_rules):
    """从规则文件列顺序推导输出字段与主/子类别列名。"""
    cols = [c for c in df_rules.columns if c not in OUTPUT_FILTER_COLUMNS]
    main_col = next((c for c in cols if "主类" in c), None)
    sub_col = next((c for c in cols if "子类" in c), None)
    return cols, main_col, sub_col


def load_bill_files(files):
    """读取多个账单文件，返回成功读取的 DataFrame 列表。"""
    dfs = []
    for f in files:
        d = sniff_and_load_data(Path(f))
        if d is not None:
            dfs.append(d)
    return dfs


def validate_final_data(df_final, main_col, sub_col):
    """输出导入前的关键字段检查。"""
    checks = [
        ('支出/收入/转账', df_final['记录类型'].isin(['支出',
         '收入', '转入', '转出']), [main_col, sub_col]),
        ('应收/应付', df_final['记录类型'].isin(['应收款项', '应付款项']),
         [main_col, sub_col, '对象']),
        ('需账户记录', df_final['记录类型'].isin(
            ['支出', '收入', '转入', '转出', '应收款项', '应付款项', '返利回馈']), ['账户']),
        ('所有记录', pd.Series(True, index=df_final.index), ['日期', '金额', '记录类型'])
    ]

    has_issues = False
    for name, mask, check_cols in checks:
        existing_cols = [c for c in check_cols if c in df_final.columns]
        missing_cols = [c for c in check_cols if c not in df_final.columns]
        if not existing_cols:
            bad = mask
        else:
            bad = mask & df_final[existing_cols].apply(
                is_blank_series).any(axis=1)
            if '金额' in existing_cols:
                bad |= mask & pd.to_numeric(
                    df_final['金额'], errors='coerce').isna()
        if missing_cols:
            bad |= mask
        if bad.any():
            has_issues = True
            missing_text = f"（缺少列: {', '.join(missing_cols)}）" if missing_cols else ""
            print(
                f"{BColors.FAIL}[!] {bad.sum()} 条【{name}】记录缺失关键信息{missing_text}:{BColors.ENDC}")
            display_cols = list(dict.fromkeys(
                c for c in ['日期', '商家', '金额', '描述'] + check_cols
                if c in df_final.columns
            ))
            print(df_final.loc[bad, display_cols].fillna(
                '').head(5).to_string(index=True))

    amount = pd.to_numeric(df_final.get(
        '金额', pd.Series(index=df_final.index)), errors='coerce')
    sign_checks = [
        ('支出/应收/转出金额应为负数',
         df_final['记录类型'].isin(['支出', '应收款项', '转出']) & (amount > 0)),
        ('收入/应付/转入金额应为正数',
         df_final['记录类型'].isin(['收入', '应付款项', '转入', '返利回馈']) & (amount < 0)),
    ]
    for name, bad in sign_checks:
        if bad.any():
            has_issues = True
            print(f"{BColors.FAIL}[!] {bad.sum()} 条【{name}】:{BColors.ENDC}")
            display_cols = [c for c in ['日期', '商家', '金额', '描述', '记录类型', main_col, sub_col, '对象']
                            if c in df_final.columns]
            print(df_final.loc[bad, display_cols].fillna(
                '').head(5).to_string(index=True))

    if not has_issues:
        print(f"{BColors.OKGREEN}✅ 数据验证通过{BColors.ENDC}")


def process_main(df_in, df_rules, main_col, sub_col):
    """处理非转账主交易。"""
    df = prepare_raw_transactions(df_in)
    df = infer_inout_from_memo(df)

    if df.empty:
        return pd.DataFrame()

    logger.info(f"分析主交易 ({len(df)} 条)")

    df = ensure_columns(df, main_col, sub_col)
    df['支付方式'] = normalize_text_series(df['支付方式']).replace(
        STANDARDIZE_ACCOUNTS, regex=True)

    df = apply_rules(df, df_rules, main_col, sub_col)
    df = apply_pdd_inout_rule(df, main_col, sub_col)
    df = mark_taobao_orders(df)
    df = construct_description(df)
    df.drop(columns=['描述_rule'], inplace=True, errors='ignore')

    df['记录类型'] = df.get('记录类型', pd.NA).replace("", pd.NA).fillna(df['收/支'])
    df['商家'] = df.get('商家', pd.NA).replace("", pd.NA).fillna(df['交易对方'])
    df = clean_auto_descriptions(df)
    df[sub_col] = normalize_text_series(df[sub_col])

    df = parse_memo_subcategory(df, main_col, sub_col)
    df = process_heuristics(df, main_col, sub_col)
    df = drop_invalid_dates(df)
    df = filter_exportable_transactions(df)
    if df.empty:
        return pd.DataFrame()

    df = finalize_records(df, main_col, sub_col)
    return df


def filter_exportable_transactions(df):
    """导出前过滤不可入账记录；原始流水先全量参与规则分析。"""
    status = normalize_text_series(df['当前状态'])
    memo = normalize_text_series(df['备注'])
    refund_text = normalize_text_series(df['当前状态']) + ' ' + memo + ' ' + \
        normalize_text_series(df.get('商品', pd.Series('', index=df.index)))
    mask_invalid_status = (
        status.isin(INVALID_STATUSES)
        | status.str.contains(FULL_REFUND_PATTERN, regex=True, na=False)
    )
    mask_zero_amount = df['金额'].abs() <= 0.0001
    mask_partial_refund_income = (
        (df['收/支'] == '收入')
        & refund_text.str.contains(PARTIAL_REFUND_PATTERN, regex=True, na=False)
    )
    mask_wechat_transfer_income = get_wechat_transfer_income_mask(df)
    mask_has_special_keyword = memo.str.contains(ALL_SPECIAL_PATTERN, na=False)
    mask_income_with_keyword = (
        (df['收/支'] == '收入')
        & memo.str.contains(INFER_INCOME_PATTERN, na=False)
    )
    mapped_record_type = normalize_text_series(
        df.get('记录类型', pd.Series('', index=df.index))
    )
    mask_income_with_rule = (
        (df['收/支'] == '收入')
        & mapped_record_type.isin(['收入', '返利回馈', '应付款项'])
    )
    keep = (
        (df['收/支'] == '支出')
        | mask_income_with_keyword
        | mask_has_special_keyword
        | mask_income_with_rule
    )
    return df[
        keep
        & ~mask_invalid_status
        & ~mask_zero_amount
        & ~mask_partial_refund_income
        & ~mask_wechat_transfer_income
    ].copy()


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
    """保存结果 CSV。"""
    if '标签' in df.columns and '_source_tag' in df.columns:
        m = ~df['记录类型'].isin(['转入', '转出'])
        df.loc[m, '标签'] = (df.loc[m, '标签'].fillna(
            "") + " " + df.loc[m, '_source_tag']).str.strip()

    type_priority = {'转出': 0, '转入': 1}
    df['_Type_Rank'] = df['记录类型'].map(type_priority).fillna(0)
    df.sort_values(['_Sort_Date', '_Type_Rank'],
                   ascending=[False, True], inplace=True)

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

    try:
        logger.info(f"项目目录: {CURRENT_DIR}")
        logger.info(f"规则文件: {RULE_BOOK_PATH}")
        logger.info(f"输出目录: {TARGET_DIR}")

        load_settings(RULE_BOOK_PATH)
        df_rules = load_rules(RULE_BOOK_PATH)
        if df_rules is None:
            if PAUSE_ON_EXIT:
                input("按 Enter 退出")
            return

        cols, main_col, sub_col = get_output_columns(df_rules)
        if not main_col or not sub_col:
            logger.error("规则文件中未找到主类别或子类别列")
            return

        files, st_date = get_user_input()
        if not files:
            return

        dfs = load_bill_files(files)
        if not dfs:
            logger.error("没有成功读取任何文件")
            return

        df_raw = pd.concat(dfs, ignore_index=True)
        df_raw['交易时间'] = df_raw['交易时间'].apply(robust_date_converter)
        if st_date is not None:
            df_raw = df_raw[df_raw['交易时间'] >= st_date]

        df_raw['交易对方'] = normalize_text_series(
            df_raw.get('交易对方', pd.Series('', index=df_raw.index)))
        target_list = [
            CONFIG['TRANSFER_TARGET_1'],
            CONFIG['TRANSFER_TARGET_2'],
            CONFIG['TRANSFER_TARGET_SNOWBALL']
        ]
        mask_trans = df_raw['交易对方'].isin(target_list)

        res_dfs = []
        main_candidates = df_raw[~mask_trans].copy()
        if mask_trans.any():
            transfer_df, transfer_unhandled = process_transfers(
                df_raw[mask_trans].copy(), main_col, sub_col
            )
            res_dfs.append(transfer_df)
            if not transfer_unhandled.empty:
                main_candidates = pd.concat(
                    [main_candidates, transfer_unhandled], ignore_index=True
                )
        if not main_candidates.empty:
            res_dfs.append(process_main(main_candidates,
                           df_rules, main_col, sub_col))

        res_dfs = [d for d in res_dfs if d is not None and not d.empty]
        if not res_dfs:
            logger.warning("无结果")
            return

        df_final = pd.concat(res_dfs, ignore_index=True)
        for c in cols:
            if c not in df_final.columns:
                df_final[c] = ""

        path = save_result(df_final, cols)
        print(f"\n{BColors.OKGREEN}✅ 成功! 文件: {path}{BColors.ENDC}")
        validate_final_data(df_final, main_col, sub_col)
        warn_refund_rows(df_raw)
        warn_wechat_transfer_income(df_raw)
        print(f"{BColors.OKGREEN}✅ 处理完成{BColors.ENDC}")

    except KeyboardInterrupt:
        logger.info("用户取消操作")
    except Exception:
        traceback.print_exc()


if __name__ == "__main__":
    start_time = time.time()
    main()
    print(f"\n运行结束 | 总耗时: {time.time() - start_time:.2f} 秒")

    if PAUSE_ON_EXIT:
        try:
            input("\n按 Enter 键退出...")
        except (KeyboardInterrupt, EOFError):
            pass
