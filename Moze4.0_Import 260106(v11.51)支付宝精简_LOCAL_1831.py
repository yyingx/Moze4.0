"""
Moze 导入脚本 v11.51 (Centralized Cleaning)
Created on Sun Jan 05 2026
Optimized: Tue Jan 06 2026
Optimization: Unified amount cleaning logic to 'sniff_and_load_data'
@author: TZY_YX
"""

import datetime
import re
import traceback
import tkinter as tk
import time
from tkinter import filedialog, simpledialog
from pathlib import Path
import numpy as np
import pandas as pd

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
    'KEYWORD_CHARGING': '自助服务-充电桩',
}

STANDARDIZE_ACCOUNTS = {
    r'.*4946.*': '平安银行4946',
    r'.*9579.*': '工商银行',
    r'.*3379.*': '招商银行Ⅱ',
    r'.*4826.*': '广发银行4826',
    r'.*零钱.*': '零钱3'
}

# --- 数据源定义 (保持不变) ---
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
        # --- 通用 ---
        "食材",
        
        # --- 面点主食 ---
        "馒头", "生水饺", "鲜面条", "干面条", "挂面",
        
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
    'Parking_fee': ["WF7023"],
    'DAILY_NECESSITIES': [
        "日用",
        "抽纸", "卷纸", "厨房纸", "垃圾袋", "保鲜袋", "保鲜膜", "洗衣液", "洗洁精", "牙膏",
        "洗发水", "一次性手套", "一次性杯", "棉签"
    ],
    'Clothing_Shoes_Bags': ["袜子", "内裤", "帽子", "手套", "鞋", "T恤", "裤", "外套", "修裤脚"],
    'Adult_Products': ["避孕套", "成人润滑剂", "安全套", "Condoms"],
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

TRIGGERS = set(AUTO_MAP_DICT.keys())
TRIGGERS.update([x[1] for x in INGREDIENT_PRIORITY if x[1]])
manual_names = ["水果", "饮料", "纯净水", "充电",
                "加油充电", "Software", "停车费", "日用", "正餐", "零食"]
TRIGGERS.update(manual_names)
TRIGGER_LIST = sorted(
    [w for w in TRIGGERS if w and not str(w).isdigit()], key=len, reverse=True)
KEYWORDS_REGEX = "|".join(map(re.escape, TRIGGER_LIST))
PREFIX_PATTERN = re.compile(rf"^(?:{KEYWORDS_REGEX})[. 。\s-]+(.*)")
VALID_SUBCATS = sorted(list(AUTO_MAP_DICT.keys()), key=len, reverse=True)
PREFIX_SUBCAT_PATTERN = re.compile(
    rf"^({'|'.join(map(re.escape, VALID_SUBCATS))})(?:[. 。\s]|$|-)")


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
    df['描述'] = df['描述'].str.replace(
        PREFIX_PATTERN, r'\1', regex=True).str.strip()
    mask_redundant = df['描述'].isin(TRIGGER_LIST) | (df['描述'] == "")
    if mask_redundant.any():
        df.loc[mask_redundant, '描述'] = ""
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
    except PermissionError:
        print(f"  {BColors.FAIL}失败：文件被 Excel 占用，请先关闭文件！{BColors.ENDC}")
        return None
    except Exception as e:
        print(f"  加载失败: {e}")
        return None


def sniff_and_load_data(file_path: Path):
    print(f"读取: {file_path.name}", flush=True)
    ftype, df = None, None
    
    try:
        # ================== 支付宝 (CSV) 精简版 ==================
        if file_path.suffix.lower() == '.csv':
            df = pd.read_csv(file_path, header=24, encoding='gb18030')
            ftype = "AliPay"

        # ================== 微信 (Excel) ==================
        elif file_path.suffix.lower() == '.xlsx':
            df = pd.read_excel(file_path, header=16, engine='openpyxl')
            ftype = "WeChat"
            
        if df is None: return None

        # ================== 统一清洗 ==================
        # 1. 清洗列名
        df.columns = df.columns.astype(str).str.strip().str.replace('[（(]', '(', regex=True).str.replace('[）)]', ')', regex=True)
        df.rename(columns=COLUMN_MAPPING, inplace=True)
        
        # 2. 必须包含金额列
        if '金额(元)' not in df.columns: return None

        # 3. 清洗金额数据
        df['金额(元)'] = pd.to_numeric(df['金额(元)'].astype(str).str.replace(r'[¥,]', '', regex=True).str.strip(), errors='coerce').fillna(0)
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
    t1, t2 = CONFIG['TRANSFER_TARGET_1'], CONFIG['TRANSFER_TARGET_2']
    mask_t1, mask_t2 = df["交易对方"] == t1, df["交易对方"] == t2
    mask_in, mask_out = df["收/支"] == "收入", df["收/支"] == "支出"
    res_list = []

    def create_records(mask, out_acc, in_acc, sub_val='转账'):
        if not mask.any():
            return
        base = df[mask].copy()

        # [优化点] 直接使用已清洗的 '金额' 列，不重复清洗

        # 转出
        o_rec = base.copy()
        o_rec['金额'] = base['金额'] * -1  # 直接乘 -1
        o_rec['记录类型'] = '转出'
        o_rec['账户'] = out_acc
        o_rec[sub_col] = sub_val

        # 转入
        i_rec = base.copy()
        # i_rec['金额'] 已经是正数，无需处理
        i_rec['记录类型'] = '转入'
        i_rec['账户'] = in_acc
        i_rec[sub_col] = sub_val

        res_list.extend([o_rec, i_rec])

    create_records(mask_t1 & mask_in, '零钱2', '零钱3')
    create_records(mask_t1 & mask_out, '零钱3', '零钱2')
    create_records(mask_t2, '零钱3', '工商银行', '提现')
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
    df = ensure_columns(df, main_col, sub_col)
    cols_to_str = [sub_col, '项目', '描述', '商家', '商品']
    df[cols_to_str] = df[cols_to_str].astype(str).apply(
        lambda x: x.str.strip().replace("nan", ""))

    # [Fix] 逐行相加
    search_series = (
        df.get('商家(old)', '').astype(str) + " " +
        df['商家'].astype(str) + " " +
        df['商品'].astype(str) + " " +
        df['描述'].astype(str)
    )

    uncat = df[sub_col] == ""
    mask = df['商品'].str.contains(CONFIG['KEYWORD_CHARGING'], na=False) & uncat
    if mask.any():
        df.loc[mask, ['名称', sub_col]] = ['充电', '加油充电']
    mask = search_series.str.contains(PATTERNS['SOFTWARE'], regex=True) & uncat
    if mask.any():
        df.loc[mask, sub_col] = 'Software'

    for k, n, s in [('FRUIT', '水果', '饮料水果'), ('DRINK', '饮料', '饮料水果'), ('WATER', '', '纯净水')]:
        mask = search_series.str.contains(PATTERNS[k], regex=True) & uncat & (
            (df[sub_col] == "") | (df[sub_col] == "食材"))
        if mask.any():
            df.loc[mask, ['名称', sub_col]] = [n, s]

    mask_meal = search_series.str.contains(PATTERNS['MEAL'], regex=True)
    main_filter = uncat | (df[main_col].isin(['购物', '居家', '饮食']))
    for key, name, sub_c in INGREDIENT_PRIORITY:
        pat = PATTERNS[key]
        mask = search_series.str.contains(
            pat, regex=True) & (~mask_meal) & main_filter
        if key in ['SEAFOOD', 'PORK', 'POULTRY', 'BEEF_MUTTON', 'VEGETABLE']:
            mask &= (~search_series.str.contains(
                PATTERNS['COOKED'], regex=True))
        if mask.any():
            df.loc[mask, [sub_col, '名称']] = [sub_c, name]

    mask = search_series.str.contains(PATTERNS['Parking_fee'], regex=True)
    if mask.any():
        df.loc[mask, [sub_col, '名称', '对象']] = ['报账', '停车费', '天之逸']

    keys = ['报账', '借出', '代付', '押金', '借入']
    debt_pat = rf"({'|'.join(keys)})\s*(.*)"
    extracted = df['描述'].str.extract(debt_pat, expand=True)
    mask_found = extracted[0].notna()
    if mask_found.any():
        df.loc[mask_found, sub_col] = extracted[0]
        mask_obj = mask_found & (extracted[1].str.strip() != "")
        if mask_obj.any():
            df.loc[mask_obj, '对象'] = extracted[1].str.strip()
        df.loc[mask_found, ['项目', '描述']] = ""

    uncat = df[sub_col] == ""
    if not uncat.empty and VALID_SUBCATS:
        extracted_series = df.loc[uncat, '描述'].str.extract(
            PREFIX_SUBCAT_PATTERN, expand=False)
        mask_hit = extracted_series.notna()
        if mask_hit.any():
            df.loc[uncat[uncat].index[mask_hit],
                   sub_col] = extracted_series[mask_hit]

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


def process_main(df, df_rules, main_col, sub_col):
    # [优化点] 不再将 '金额' 初始化为空字符串，防止后续类型错误
    for c in ['当前状态', '收/支', '交易对方', '交易时间']:
        if c not in df.columns:
            df[c] = ""

    # [优化点] 过滤时使用已存在的数字 '金额' 列
    df = df[
        (~df["当前状态"].isin(["已全额退款", "交易关闭"])) &
        (df["收/支"] == "支出") &
        (abs(df["金额"]) > 0.0001)  # 使用 abs 防止精度问题，确保不为0
    ].copy()

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

    if '商户单号' in df.columns:
        mask_tb = (df.get('_source_tag', '') == '#AliPay') & df['商户单号'].astype(
            str).str.strip().str.startswith('T200P4', na=False)
        if mask_tb.any():
            df.loc[mask_tb, '商家'] = '淘宝'

    df = construct_description(df)
    if '描述_rule' in df.columns:
        df.drop(columns=['描述_rule'], inplace=True)
    df['记录类型'] = df.get('记录类型', pd.NA).replace("", pd.NA).fillna(df['收/支'])
    df['商家'] = df.get('商家', pd.NA).replace("", pd.NA).fillna(df['交易对方'])

    df[sub_col] = df[sub_col].astype(str).str.strip().replace("nan", "")

    # [Fix] 逐行拼接
    search_series = (
        df.get('商家(old)', '').astype(str) + " " +  # <--- 核心修改：加上这一行
        df['商家'].astype(str) + " " +
        df['商品'].astype(str) + " " +
        df['描述'].astype(str)
    )

    mask_time = (df[sub_col] == "") & search_series.str.contains(
        PATTERNS['MEAL'], regex=True)
    if mask_time.any():
        h = df.loc[mask_time, '交易时间'].dt.hour
        conditions = [(h >= 6) & (h < 11), (h >= 11) &
                      (h < 16), (h >= 16) & (h < 21)]
        choices = ["早餐", "午餐", "晚餐"]
        df.loc[mask_time, sub_col] = np.select(
            conditions, choices, default="夜宵")

    df = process_heuristics(df, main_col, sub_col)

    # [优化点] 不再重复计算 df['金额']，只处理负数
    mask_neg = (df['记录类型'].isin(['支出', '应付款项', '应收款项'])) & (
        df.get('收/支') == '支出')
    df.loc[mask_neg, '金额'] *= -1

    mask_borrow = df[sub_col] == '借入'
    df.loc[mask_borrow, '金额'] = df.loc[mask_borrow, '金额'].abs()

    mask_fix_pay = (df['支付方式'] == "/") & (df['记录类型'] == "收入")
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
    print(f"{BColors.BOLD}=== Moze 导入脚本 v11.51 (Optimized Clean) ==={BColors.ENDC}")
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
        df_raw['交易对方'] = df_raw.get('交易对方', pd.NA).astype(str).str.strip()
        mask_trans = df_raw['交易对方'].isin(
            [CONFIG['TRANSFER_TARGET_1'], CONFIG['TRANSFER_TARGET_2']])
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
        checks = [('支出', df_final['记录类型'] == '支出', [main_col, sub_col, '项目']),
                  ('债权', df_final['记录类型'].isin(['应收款项', '应付款项']), [main_col, sub_col, '对象'])]
        for name, mask, check_cols in checks:
            bad = mask & df_final[check_cols].isin(["", pd.NA]).any(axis=1)
            if bad.any():
                print(f"{BColors.FAIL}[严重] {bad.sum()} 条【{
                      name}】记录缺失关键信息:{BColors.ENDC}")
                # 定义要显示的列：基础信息 + 检查的列(主类/子类/项目)
                display_cols = ['日期', '商家', '金额', '描述'] + check_cols
                # 打印前5条
                print(df_final.loc[bad, display_cols].fillna(
                    '[空]').head(5).to_string(index=True))
                print(f"{BColors.OKGREEN}√ {name}记录完美{BColors.ENDC}")
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
