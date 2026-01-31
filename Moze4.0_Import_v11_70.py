# -*- coding: utf-8 -*-
"""
Moze 导入脚本 v11.70 (JSON Config)
Created on Sun Jan 05 2026
Optimized: Thu Jan 30 2026

@author: TZY_YX

CHANGELOG (v11.70):
[新增] 分离 INGREDIENT_CONFIG 到 JSON 文件，便于维护
[新增] 分离 RAW_MAPPING_CONFIG 到 JSON 文件
[新增] 支持从 ingredient_config.json 加载分类配置
[保留] v11.69 所有功能

配置文件: ingredient_config.json
- INGREDIENT_CONFIG: 分类关键词配置
- MEAL_KEYWORDS: 正餐关键词
- PARKING_KEYWORDS: 停车费关键词
- MEAL_TIME_RANGES: 餐食时间段
- RAW_MAPPING_CONFIG: 子类别到记录类型/主类别/项目的映射
"""

# === 版本信息 ===
__version__ = '11.70'
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
INGREDIENT_CONFIG_PATH = CURRENT_DIR / "ingredient_config.json"

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

# 名称关键词映射
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
#    [JSON] 从配置文件加载
# ==========================================

def load_ingredient_config(config_path: Path):
    """从 JSON 文件加载分类配置"""
    if not config_path.exists():
        logger.error(f"配置文件不存在: {config_path}")
        return None, None, None, None, None

    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        ingredient_config = data.get('INGREDIENT_CONFIG', {})
        meal_keywords = data.get('MEAL_KEYWORDS', [])
        parking_keywords = data.get('PARKING_KEYWORDS', [])
        meal_time_ranges = data.get('MEAL_TIME_RANGES', {})
        raw_mapping_list = data.get('RAW_MAPPING_CONFIG', [])

        # 转换 MEAL_TIME_RANGES 的列表为元组
        meal_time_ranges = {k: tuple(v) for k, v in meal_time_ranges.items()}

        # 转换 RAW_MAPPING_CONFIG 从数组格式到字典格式
        raw_mapping_config = {}
        for item in raw_mapping_list:
            key = (item['record_type'], item['main_category'], item['project'])
            raw_mapping_config[key] = item['subcategories']

        logger.info(f"已加载配置: {len(ingredient_config)} 个分类, {len(meal_keywords)} 个正餐关键词, {len(raw_mapping_config)} 个映射规则")
        return ingredient_config, meal_keywords, parking_keywords, meal_time_ranges, raw_mapping_config

    except json.JSONDecodeError as e:
        logger.error(f"JSON 解析错误: {e}")
    except Exception as e:
        logger.error(f"加载配置失败: {type(e).__name__} - {e}")

    return None, None, None, None, None


# 加载配置
INGREDIENT_CONFIG, MEAL_KEYWORDS, PARKING_KEYWORDS, MEAL_TIME_RANGES, RAW_MAPPING_CONFIG = load_ingredient_config(INGREDIENT_CONFIG_PATH)

# 如果加载失败，使用默认空配置
if INGREDIENT_CONFIG is None:
    logger.warning("使用空配置，分类功能将不可用")
    INGREDIENT_CONFIG = {}
    MEAL_KEYWORDS = []
    PARKING_KEYWORDS = []
    MEAL_TIME_RANGES = {'早餐': (6, 11), '午餐': (11, 16), '晚餐': (16, 21)}
    RAW_MAPPING_CONFIG = {}


# ==========================================
#    [构建] 编译后的数据结构
# ==========================================

def build_data_structures():
    """从 INGREDIENT_CONFIG 构建运行时数据结构"""
    data_source = {}
    patterns = {}

    for key, config in INGREDIENT_CONFIG.items():
        data_source[key] = config['keywords']
        patterns[key] = re.compile(r"(?:" + "|".join(map(re.escape, config['keywords'])) + ")")

    # 添加正餐和停车费
    if MEAL_KEYWORDS:
        data_source['MEAL'] = MEAL_KEYWORDS
        patterns['MEAL'] = re.compile(r"(?:" + "|".join(map(re.escape, MEAL_KEYWORDS)) + ")")

    if PARKING_KEYWORDS:
        data_source['Parking_fee'] = PARKING_KEYWORDS
        patterns['Parking_fee'] = re.compile(r"(?:" + "|".join(map(re.escape, PARKING_KEYWORDS)) + ")")

    # 添加报销关键词
    data_source['REIM_TRAVEL'] = REIM_TRAVEL_KEYS
    data_source['REIM_EXPENSE'] = REIM_EXPENSE_KEYS

    return data_source, patterns


# 构建运行时结构
DATA_SOURCE, PATTERNS = build_data_structures()

# 按优先级排序的配置列表
SORTED_INGREDIENT_CONFIG = sorted(
    INGREDIENT_CONFIG.items(),
    key=lambda x: x[1].get('priority', 999)
)

# RAW_MAPPING_CONFIG 从 JSON 配置文件加载

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
                df = pd.read_csv(file_path, header=i, encoding=encoding, nrows=1)
            else:
                df = pd.read_excel(file_path, header=i, engine='openpyxl', nrows=1)

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
                logger.error(f"规则文件中缺少 'Moze Dict' sheet")
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

        df['金额(元)'] = pd.to_numeric(
            df['金额(元)'].astype(str).str.replace(r'[¥,]', '', regex=True).str.strip(),
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
    """处理报销关键词"""
    df = df.copy()
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
    """处理通用分类词"""
    df = df.copy()
    generic_extracted = df['描述'].str.extract(GENERIC_PATTERN, expand=True)
    mask_generic = generic_extracted[0].notna() & (df[sub_col] == "")

    if not mask_generic.any():
        return df

    matched_keys = generic_extracted.loc[mask_generic, 0]
    tails = generic_extracted.loc[mask_generic, 1].str.strip()

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

    for keyword, (subcat, name) in NAME_KEYWORDS.items():
        mask_kw = matched_keys == keyword
        if mask_kw.any():
            kw_idx = mask_kw[mask_kw].index
            df.loc[kw_idx, sub_col] = subcat
            df.loc[kw_idx, '名称'] = name
            df.loc[kw_idx, '描述'] = tails.loc[kw_idx]

    return df


def process_heuristics(df_in, main_col, sub_col):
    """启发式分类推导"""
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

    uncat = df[sub_col] == ""
    mask_meal = search_series.str.contains(PATTERNS.get('MEAL', re.compile('')), regex=True) if 'MEAL' in PATTERNS else pd.Series(False, index=df.index)
    mask_ingredients_exact = search_series.str.contains(
        PATTERNS.get('INGREDIENTS', re.compile('')), regex=True) if 'INGREDIENTS' in PATTERNS else pd.Series(False, index=df.index)

    mask_cooked = search_series.str.contains(PATTERNS.get('COOKED', re.compile('')), regex=True) if 'COOKED' in PATTERNS else pd.Series(False, index=df.index)

    main_filter = uncat | (df[main_col].isin(['购物', '居家', '饮食']))

    for key, config in tqdm(SORTED_INGREDIENT_CONFIG, desc="分类推导", leave=False, disable=not HAS_TQDM):
        if key not in PATTERNS:
            continue

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

        if 'COOKED' in exclude_if_match:
            mask &= (~mask_cooked)

        if mask.any():
            df.loc[mask, [sub_col, '名称']] = [sub_c, name]
            if obj:
                df.loc[mask, '对象'] = obj

    # 停车费
    if 'Parking_fee' in PATTERNS:
        mask = search_series.str.contains(PATTERNS['Parking_fee'], regex=True)
        if mask.any():
            df.loc[mask, [sub_col, '名称', '对象']] = ['报账', '停车费', '天之逸']

    df = process_debt_keywords(df, sub_col)
    df = process_reimbursement(df, sub_col)
    df = process_generic_keywords(df, sub_col)

    # MEAL 时间段分类
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


def parse_memo_subcategory(df, main_col, sub_col):
    """解析备注中的子类别"""
    if '备注' not in df.columns:
        return df

    df = df.copy()
    df['_is_meal_from_memo'] = False
    memo_series = df['备注'].astype(str).str.strip().replace('nan', '')

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

    pattern = rf'^正餐(.*)$'
    matches = memo_series.str.match(pattern, na=False)
    if matches.any():
        extracted = memo_series[matches].str.replace(pattern, r'\1', regex=True).str.strip()
        df.loc[matches, '描述'] = extracted
        df.loc[matches, '_is_meal_from_memo'] = True
        memo_series = memo_series.where(~matches, '')

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

        if subcat in RECEIVABLE_PAYABLE_SUBCATS:
            matched_indices = matches[matches].index
            mask_dot = extracted.str.contains(r'\.', regex=True, na=False)

            dot_indices = matched_indices[mask_dot[matched_indices]]
            if len(dot_indices) > 0:
                dot_content = extracted.loc[dot_indices]
                split_df = dot_content.str.split('.', n=1, expand=True)

                obj_values = split_df[0].str.strip()
                desc_values = split_df[1].str.strip() if 1 in split_df.columns else pd.Series('', index=dot_indices)

                empty_obj_mask = (obj_values == '') | obj_values.isna()
                if empty_obj_mask.any():
                    obj_values = obj_values.where(
                        ~empty_obj_mask,
                        dot_content.str.replace('.', '', regex=False).str.strip()
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
    """应用字典规则匹配"""
    df = df.copy()
    df['商家(old)'] = df.get('交易对方', pd.NA).astype(str).str.strip()
    rename_map = {c: f'{c}_rule' for c in df_rules.columns if c not in ['商家(old)', 'is_regex']}

    df = pd.merge(
        df,
        df_rules[df_rules['is_regex'] == 0].rename(columns=rename_map).drop(columns='is_regex', errors='ignore'),
        on='商家(old)', how='left'
    )
    df.reset_index(drop=True, inplace=True)

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

    for orig, rule in rename_map.items():
        if rule in df.columns and orig != '描述':
            df[orig] = df[rule].combine_first(df[orig])
            df.drop(columns=[rule], inplace=True)

    return df


def finalize_records(df, main_col, sub_col):
    """最终处理"""
    df = df.copy()

    mask_neg = (df['记录类型'].isin(['支出', '应付款项', '应收款项'])) & (df.get('收/支') == '支出')
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

    df = df[
        (~df["当前状态"].isin(["已全额退款", "交易关闭"])) &
        (abs(df["金额"]) > 0.0001)
    ].copy()

    all_reim_keywords = REIM_TRAVEL_KEYS + REIM_EXPENSE_KEYS
    reim_pattern = '|'.join(all_reim_keywords)

    if '备注' in df.columns:
        memo = df['备注'].astype(str).str.strip()
        mask_empty_inout = df['收/支'].fillna('').astype(str).str.strip().isin(['', 'nan', 'NaN', 'None'])

        mask_borrow_in = mask_empty_inout & memo.str.contains(r'^借入', na=False, regex=True)
        if mask_borrow_in.any():
            df.loc[mask_borrow_in, '收/支'] = '收入'

        mask_borrow_out = mask_empty_inout & memo.str.contains(r'^(?:借出|代付|报账|押金)', na=False, regex=True)
        if mask_borrow_out.any():
            df.loc[mask_borrow_out, '收/支'] = '支出'

        mask_reim = mask_empty_inout & memo.str.contains(rf'^(?:{reim_pattern})', na=False, regex=True)
        if mask_reim.any():
            df.loc[mask_reim, '收/支'] = '支出'

    memo_series = df['备注'].astype(str).str.strip()
    all_special_pattern = '|'.join(DEBT_KEYWORDS) + '|' + reim_pattern
    mask_has_special_keyword = memo_series.str.contains(rf'^(?:{all_special_pattern})', na=False, regex=True)
    df = df[(df["收/支"] == "支出") | mask_has_special_keyword].copy()

    if df.empty:
        return pd.DataFrame()

    logger.info(f"处理主交易 ({len(df)} 条)")

    df = ensure_columns(df, main_col, sub_col)
    df['支付方式'] = df['支付方式'].astype(str).replace(STANDARDIZE_ACCOUNTS, regex=True)

    df = apply_rules(df, df_rules, main_col, sub_col)

    if '商户单号' in df.columns:
        mask_tb = (df.get('_source_tag', '') == '#AliPay') & df['商户单号'].astype(str).str.strip().str.startswith('T200P', na=False)
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
    df = df.copy()

    if '标签' in df.columns:
        m = ~df['记录类型'].isin(['转入', '转出'])
        df.loc[m, '标签'] = (df.loc[m, '标签'].fillna("") + " " + df.loc[m, '_source_tag']).str.strip()

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
    print(f"  Moze 导入脚本 v{__version__} (JSON Config)")
    print(f"  作者: {__author__} | 更新: {__updated__}")
    print("=" * 50)
    print(f"{BColors.ENDC}")

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
                print(df_final.loc[bad, display_cols].fillna('').head(5).to_string(index=True))

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
