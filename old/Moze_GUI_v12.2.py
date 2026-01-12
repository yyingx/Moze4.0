"""
Moze 导入脚本 v12.2 (GUI + 最终校验版)
核心逻辑：v11.39 + 自动质检报告
界面框架：FreeSimpleGUI / PySimpleGUI
Created on Mon Jan 06 2026
@author: TZY_YX
"""

import FreeSimpleGUI as sg
import pandas as pd
import numpy as np
import datetime
import re
import traceback
import time
from pathlib import Path
import warnings

# ==========================================
#      1. 配置与常量 (保持 v11.39 原样)
# ==========================================


class BColors:
    OKGREEN = ''
    WARNING = ''
    FAIL = ''
    ENDC = ''
    BOLD = ''


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

DATA_SOURCE = {
    'FRUIT': [
        "水果", "果园", "果蔬", "百果园", "鲜果", "苹果", "香蕉", "红枣", "大枣", "桃子",
        "草莓", "西瓜", "柠檬", "橙子", "榴莲", "芒果", "车厘子", "葡萄"
    ],
    'DRINK': ["饮料", "可乐", "红牛", "奶茶", "咖啡", "拿铁", "酸奶", "椰汁"],
    'WATER': ["纯净水", "矿泉水", "农夫山泉", "怡宝", "百岁山"],
    'SOFTWARE': ["软件", "APP", "应用", "安卓"],
    'CHARGING': ["特来电", "星星充电", "小桔充电", "国家电网", "电费", "充电", "蔚来", "特斯拉"],
    'VEGETABLE': [
        "蔬菜", "白菜", "菠菜", "生菜", "土豆", "萝卜", "红薯", "西红柿", "黄瓜",
        "辣椒", "玉米", "香菇", "金针菇", "大葱", "大蒜"
    ],
    'Snack': ["零食", "蛋糕", "面包", "巧克力", "瓜子", "坚果"],
    'BEAN_PRODUCT': ["豆腐", "豆皮", "腐竹", "豆干"],
    'INGREDIENTS': ["食材", "馒头", "火腿", "老干妈", "酱料", "盐", "油", "米", "面", "挂面"],
    'PORK': ["猪肉", "排骨", "五花肉", "瘦肉", "猪蹄"],
    'POULTRY': ["鸡肉", "鸡腿", "鸡翅", "鸭肉", "鸭腿"],
    'BEEF_MUTTON': ["牛肉", "牛排", "肥牛", "羊肉"],
    'SEAFOOD': ["鱼", "虾", "蟹", "海鲜", "生蚝"],
    'Eggs': ["鸡蛋", "皮蛋", "咸鸭蛋"],
    'COOKED': ["熟食", "卤菜", "凉菜", "烤鸭", "烧鸡", "肉丸"],
    'RICE': ["大米"],
    'MEAL': [
        "正餐", "食堂", "外勤", "麦当劳", "肯德基", "必胜客", "老乡鸡", "麦香园", "沙县",
        "兰州", "麻辣烫", "火锅", "烧烤", "炒饭", "盖浇饭", "面馆", "早餐", "包子", "饺子"
    ],
    'Parking_fee': ["WF7023", "停车"],
    'DAILY_NECESSITIES': ["日用", "纸巾", "洗衣液", "洗发水", "牙膏", "垃圾袋"],
    'Clothing_Shoes_Bags': ["衣服", "裤子", "鞋", "袜子", "外套"],
    'Adult_Products': ["避孕套", "安全套"],
    'SERVER': ["节点", "Dler"],
    'Furniture_HomeTextiles': ["被子", "枕头", "床单"]
}

INGREDIENT_PRIORITY = [
    ('VEGETABLE', '蔬菜', '食材'), ('RICE', '大米', '食材'), ('BEAN_PRODUCT', '豆制品', '食材'),
    ('PORK', '猪肉', '食材'), ('BEEF_MUTTON', '牛羊肉', '食材'), ('POULTRY', '禽肉', '食材'),
    ('SEAFOOD', '海鲜水产', '食材'), ('Eggs', '蛋及蛋制品', '食材'), ('INGREDIENTS', '', '食材'),
    ('COOKED', '熟食', '食材'), ('Snack', '', '零食'), ('DAILY_NECESSITIES', '', '日常用品'),
    ('Clothing_Shoes_Bags', '', '服饰鞋包'), ('Furniture_HomeTextiles', '', '家具家纺'),
    ('SERVER', '节点', '虚拟其他'), ('Adult_Products', 'Condoms', '保健用品')
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

PATTERNS = {k: r"(?:" + "|".join(map(re.escape, v)) +
            ")" for k, v in DATA_SOURCE.items()}

AUTO_MAP_DICT = {}
for (r_type, main_cat, proj), sub_list in RAW_MAPPING_CONFIG.items():
    for sub in sub_list:
        AUTO_MAP_DICT[sub] = (r_type, main_cat, proj)

COLUMN_MAPPING = {
    '交易时间': '交易时间', '交易创建时间': '交易时间', '付款时间': '交易时间',
    '交易类型': '交易类型', '交易分类': '交易类型',
    '交易对方': '交易对方', '商品': '商品', '商品名称': '商品', '商品说明': '商品',
    '收/支': '收/支', '金额(元)': '金额(元)', '金额（元）': '金额(元)', '金额': '金额(元)',
    '支付方式': '支付方式', '收/付款方式': '支付方式',
    '当前状态': '当前状态', '交易状态': '当前状态',
    '备注': '备注',
    '交易单号': '交易单号', '交易订单号': '交易单号', '商户单号': '商户单号', '商家订单号': '商户单号'
}

# ==========================================
#      2. 业务逻辑函数 (v11.39)
# ==========================================


def load_settings(rule_path: Path):
    if not rule_path.exists():
        return
    try:
        with pd.ExcelFile(rule_path, engine='openpyxl') as xl:
            df = pd.read_excel(xl, sheet_name='Settings')
        for _, row in df.iterrows():
            k, v = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
            if k in CONFIG and v and v.lower() != 'nan':
                CONFIG[k] = v
        print("已加载自定义配置")
    except:
        pass


def load_rules(rule_path: Path):
    """加载规则（防文件占用）"""
    print("正在加载规则...", flush=True)
    try:
        with pd.ExcelFile(rule_path, engine='openpyxl') as xl:
            s_name = next((s for s in xl.sheet_names if s.lower(
            ) != 'settings' and '商家(old)' in pd.read_excel(xl, sheet_name=s, nrows=1).columns), None)
            if not s_name:
                return None
            df = pd.read_excel(xl, sheet_name=s_name)

        df['is_regex'] = pd.to_numeric(
            df.get('is_regex', 0), errors='coerce').fillna(0)
        df['商家(old)'] = df['商家(old)'].astype(str).str.strip()
        print(f"  已加载 {len(df)} 条规则")
        return df
    except PermissionError:
        print("❌ 错误：字典文件正被 Excel 占用！请关闭文件后重试。")
        return None
    except Exception as e:
        print(f"  加载失败: {e}")
        return None


def sniff_and_load_data(file_path: Path):
    print(f"读取: {file_path.name}", flush=True)
    ftype, df = None, None
    try:
        if file_path.suffix.lower() == '.csv':
            for enc in ['gb18030', 'utf-8']:
                try:
                    preview = pd.read_csv(
                        file_path, header=None, nrows=50, encoding=enc, names=list(range(30)))
                    mask = preview.apply(lambda x: x.astype(str).str.contains(
                        '交易创建时间|交易时间|付款时间', na=False)).any(axis=1)
                    header_row = mask.idxmax() if mask.any() else 24
                    df = pd.read_csv(
                        file_path, header=header_row, encoding=enc)
                    ftype = "AliPay"
                    break
                except:
                    continue
        elif file_path.suffix.lower() == '.xlsx':
            try:
                df = pd.read_excel(file_path, header=16, engine='openpyxl')
                ftype = "WeChat"
            except:
                pass

        if df is None:
            return None
        df.columns = df.columns.astype(str).str.strip().str.replace(
            '（', '(', regex=False).str.replace('）', ')', regex=False)
        df.rename(columns=COLUMN_MAPPING, inplace=True)
        if '金额(元)' not in df.columns:
            return None
        df['金额(元)'] = pd.to_numeric(df['金额(元)'].astype(str).str.replace(
            '¥', '').str.replace(',', '').str.strip(), errors='coerce').fillna(0)
        df['_source_tag'] = '#WechatPay' if ftype == "WeChat" else '#AliPay'
        return df
    except Exception as e:
        print(f"  加载错误: {e}")
        return None


def ensure_columns(df, main_col, sub_col):
    cols = [main_col, sub_col, '金额', '记录类型', '账户', '商家', '描述',
            '项目', '名称', '对象', '标签', '日期', '时间', '币种', '手续费', '折扣']
    for c in cols:
        if c not in df.columns:
            df[c] = pd.NA
    for c in ['记录类型', '项目', '对象', main_col, sub_col, '商家', '名称']:
        df[c] = df[c].fillna("")
    return df


def process_transfers(df, main_col, sub_col):
    res = []
    t1, t2 = CONFIG['TRANSFER_TARGET_1'], CONFIG['TRANSFER_TARGET_2']

    def add_pair(sub_df, out_a, in_a, sub_val='转账'):
        if sub_df.empty:
            return
        o_rec, i_rec = sub_df.copy(), sub_df.copy()
        o_rec['金额'] = pd.to_numeric(o_rec['金额(元)']) * -1
        o_rec['记录类型'], o_rec['账户'], o_rec[sub_col] = '转出', out_a, sub_val
        i_rec['金额'] = pd.to_numeric(i_rec['金额(元)'])
        i_rec['记录类型'], i_rec['账户'], i_rec[sub_col] = '转入', in_a, sub_val
        res.extend([o_rec, i_rec])

    add_pair(df[(df["交易对方"] == t1) & (df["收/支"] == "收入")], '零钱2', '零钱3')
    add_pair(df[(df["交易对方"] == t1) & (df["收/支"] == "支出")], '零钱3', '零钱2')
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
    df = ensure_columns(df, main_col, sub_col)
    for col in [sub_col, '项目', '描述', '商家', '商品']:
        df[col] = df[col].astype(str).str.strip().replace("nan", "")

    def check(k):
        pat = PATTERNS.get(k, "^$")
        return (df['商家'].str.contains(pat, regex=True) |
                df['商品'].str.contains(pat, regex=True) |
                df['描述'].str.contains(pat, regex=True))

    uncat = df[sub_col] == ""
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
    mask = check('Parking_fee')
    if mask.any():
        df.loc[mask, [sub_col, '名称', '对象']] = ['报账', '停车费', '天之逸']

    # 债权
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

    # 前缀提取
    uncat = df[sub_col] == ""
    valid_subcats = sorted(list(AUTO_MAP_DICT.keys()), key=len, reverse=True)
    pattern_prefix = rf"^({'|'.join(
        map(re.escape, valid_subcats))})(?:[. 。\s]|$|-)"
    if not uncat.empty and valid_subcats:
        extracted_series = df.loc[uncat, '描述'].astype(
            str).str.extract(pattern_prefix, expand=False)
        mask_hit = extracted_series.notna()
        if mask_hit.any():
            idx = uncat[uncat].index[mask_hit]
            print(f"  [提示] 智能提取: {len(idx)} 条 (如: {
                  extracted_series[mask_hit].iloc[0]}.xx)")
            df.loc[idx, sub_col] = extracted_series[mask_hit]

    # 全自动推导
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


def construct_description(df):
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
    triggers = set()
    triggers.update(AUTO_MAP_DICT.keys())
    # 想要保留具体餐品名，可以注释掉下面这行
    if 'MEAL' in DATA_SOURCE:
        triggers.update(DATA_SOURCE['MEAL'])
    triggers.update([x[1] for x in INGREDIENT_PRIORITY if x[1]])
    manual_names = ["水果", "饮料", "纯净水", "充电",
                    "加油充电", "Software", "停车费", "日用", "正餐", "零食"]
    triggers.update(manual_names)
    trigger_list = [w for w in triggers if w and not str(w).isdigit()]
    trigger_list.sort(key=len, reverse=True)
    keywords_regex = "|".join(map(re.escape, trigger_list))
    prefix_pat = rf"^(?:{keywords_regex})[. 。\s-]+(.*)"
    df['描述'] = df['描述'].astype(str).str.replace(
        prefix_pat, r'\1', regex=True).str.strip()
    mask_redundant = df['描述'].isin(trigger_list) | (df['描述'] == "")
    if mask_redundant.any():
        df.loc[mask_redundant, '描述'] = ""
    return df


def robust_date_converter(x):
    if isinstance(x, (datetime.datetime, datetime.date)):
        return pd.to_datetime(x)
    try:
        if isinstance(x, (int, float)):
            return pd.to_datetime(x, unit='D', origin='1899-12-30')
        return pd.to_datetime(x, errors='coerce')
    except:
        return pd.NaT


def process_main(df, df_rules, main_col, sub_col):
    for c in ['当前状态', '收/支', '交易对方', '交易时间']:
        if c not in df.columns:
            df[c] = ""
    df = df[(~df["当前状态"].isin(["已全额退款", "交易关闭"])) & (df["收/支"] == "支出")].copy()
    if df.empty:
        return pd.DataFrame()
    print(f"--- 处理交易 ({len(df)}条) ---", flush=True)

    df = ensure_columns(df, main_col, sub_col)
    df['支付方式'] = df['支付方式'].astype(str).replace(
        STANDARDIZE_ACCOUNTS, regex=True)

    df['商家(old)'] = df.get('交易对方', pd.NA).astype(str).str.strip()
    rename_map = {c: f'{c}_rule' for c in df_rules.columns if c not in [
        '商家(old)', 'is_regex']}
    df = pd.merge(df, df_rules[df_rules['is_regex'] == 0].rename(columns=rename_map).drop(
        columns='is_regex', errors='ignore'), on='商家(old)', how='left')

    uncat = df[df[f"{main_col}_rule"].isna()].index if f"{
        main_col}_rule" in df.columns else df.index
    if not uncat.empty:
        to_match = df.loc[uncat, '交易对方'].astype(str)
        for _, r in df_rules[df_rules['is_regex'] == 1].iterrows():
            with warnings.catch_warnings():
                warnings.filterwarnings("ignore")
                idx = to_match[to_match.str.contains(
                    r['商家(old)'], regex=True)].index
            if not idx.empty:
                for c, rc in rename_map.items():
                    df.loc[idx, rc] = r[c]
                to_match = to_match.drop(idx)

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

    mask_time = (df[sub_col] == "")
    if mask_time.any():
        h = df.loc[mask_time, '交易时间'].dt.hour
        df.loc[mask_time, sub_col] = np.select(
            [(h >= 6) & (h < 11), (h >= 11) & (h < 16), (h >= 16) & (h < 21)],
            ["早餐", "午餐", "晚餐"], default="夜宵"
        )

    df = process_heuristics(df, main_col, sub_col)

    df['金额'] = pd.to_numeric(df.get('金额(元)', 0), errors='coerce').fillna(0)
    df.loc[(df['记录类型'].isin(['支出', '应付款项', '应收款项']))
           & (df.get('收/支') == '支出'), '金额'] *= -1
    df.loc[df[sub_col] == '借入', '金额'] = df.loc[df[sub_col] == '借入', '金额'].abs()
    df.loc[(df['支付方式'] == "/") & (df['记录类型'] == "收入"), '支付方式'] = "零钱3"

    df['日期'] = df['交易时间'].dt.strftime('%Y/%m/%d')
    df['时间'] = df['交易时间'].dt.strftime('%H:%M:%S')
    df['_Sort_Date'] = df['交易时间']
    df['账户'] = df.get('支付方式', "")
    df.loc[:, ['币种', '手续费', '折扣']] = ["CNY", 0, 0]
    df.loc[df['记录类型'].isin(['应收款项', '应付款项']), '商家'] = ""
    mask_debt = df['记录类型'].isin(['应收款项', '应付款项'])
    if mask_debt.any():
        df.loc[mask_debt, '项目'] = ""
    df = clean_final_description(df)
    return df


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

# ==========================================
#      3. GUI 界面逻辑 (View Controller)
# ==========================================


def processing_task(files, date_str, window):
    """任务调度器"""
    start_time = time.time()
    load_settings(RULE_BOOK_PATH)
    df_rules = load_rules(RULE_BOOK_PATH)
    if df_rules is None:
        print("❌ 错误：无法加载字典规则。")
        return

    cols = [c for c in df_rules.columns if c not in ['商家(old)', 'is_regex']]
    main_col = next((c for c in cols if "主类" in c), None)
    sub_col = next((c for c in cols if "子类" in c), None)

    dfs = [d for f in files if (d := sniff_and_load_data(Path(f))) is not None]
    if not dfs:
        print("❌ 未加载任何有效数据。")
        return

    df_raw = pd.concat(dfs, ignore_index=True)
    df_raw['交易时间'] = df_raw['交易时间'].apply(robust_date_converter)

    if date_str:
        try:
            st_date = pd.to_datetime(date_str, format="%Y%m%d")
            print(f"筛选日期: {date_str} 之后")
            df_raw = df_raw[df_raw['交易时间'] >= st_date]
        except:
            print("⚠️ 日期格式错误，已忽略筛选。")

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
        print("⚠️ 处理后无结果 (可能所有记录都被筛选掉了)")
        return

    df_final = pd.concat(res_dfs, ignore_index=True)
    for c in cols:
        if c not in df_final.columns:
            df_final[c] = ""

    # ----------------------------------------------------
    #  [核心新增] 最终数据校验 & 统计
    # ----------------------------------------------------
    print("\n🔍 正在进行最终校验...")
    checks = [
        ('支出', df_final['记录类型'] == '支出', [main_col, sub_col, '项目']),
        ('债权', df_final['记录类型'].isin(
            ['应收款项', '应付款项']), [main_col, sub_col, '对象'])
    ]

    total_bad = 0
    for name, mask, check_cols in checks:
        # 找出关键列为空的行
        bad_mask = mask & df_final[check_cols].isin(
            ["", pd.NA, "nan"]).any(axis=1)
        bad_count = bad_mask.sum()

        if bad_count > 0:
            total_bad += bad_count
            print(f"❌ [严重] 发现 {bad_count} 条【{
                  name}】记录缺失关键信息 ({'/'.join(check_cols)}):")
            # 打印前5条有问题的记录，方便检查
            print(df_final.loc[bad_mask, ['日期', '商家', '描述',
                  '金额'] + check_cols].head(5).to_string())
            print("...")
        else:
            print(f"✅ {name}记录完整。")

    if total_bad == 0:
        print("✨ 所有记录均符合要求！")
    else:
        print(f"⚠️ 总计 {total_bad} 条记录不符合要求，请检查生成的 Excel 文件。")
    # ----------------------------------------------------

    path = save_result(df_final, cols)
    print("--------------------------------------------------")
    print(f"✅ 导出完成! \n文件: {path}")
    print(f"耗时: {time.time() - start_time:.2f} 秒")
    print("--------------------------------------------------")


def main_gui():
    sg.theme('SystemDefault')
    layout = [
        [sg.Text("Moze 导入工具 v12.2", font=("Helvetica", 16), text_color="blue")],
        [sg.Text("选择账单文件:", size=(15, 1)),
         sg.Input(key='-FILES-', enable_events=True,
                  readonly=True, size=(50, 1)),
         sg.FilesBrowse("浏览...", file_types=(("账单文件", "*.csv;*.xlsx"),))],
        [sg.Text("起始日期:", size=(15, 1)),
         sg.Input(key='-DATE-', size=(20, 1), tooltip="格式 YYYYMMDD"),
         sg.Text("(格式: 20260101)", text_color="gray")],
        [sg.Button("🚀 开始导入", size=(15, 1), button_color="green"),
         sg.Button("退出", size=(10, 1))],
        [sg.HorizontalSeparator()],
        [sg.Output(size=(85, 20), key='-OUTPUT-', font=("Consolas", 10))]
    ]
    window = sg.Window("Moze Automation", layout)
    while True:
        event, values = window.read()
        if event in (sg.WINDOW_CLOSED, "退出"):
            break
        if event == "🚀 开始导入":
            file_str = values['-FILES-']
            if not file_str:
                print("❌ 请先选择文件！")
                continue
            files = file_str.split(";")
            date_str = values['-DATE-']
            print(f"\n🚀 开始处理 {len(files)} 个文件...")
            processing_task(files, date_str, window)
    window.close()


if __name__ == "__main__":
    main_gui()
