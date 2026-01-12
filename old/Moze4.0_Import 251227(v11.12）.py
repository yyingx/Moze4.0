import pandas as pd
import os
import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog
import openpyxl
import re
import numpy as np
import traceback
from pathlib import Path

# ==========================================
#      Moze 导入脚本 v11.12 (完整优化版)
# ==========================================

class bcolors:
    OKGREEN = '\033[92m'; WARNING = '\033[93m'; FAIL = '\033[91m'; ENDC = '\033[0m'; BOLD = '\033[1m'

# --- [配置 1] 路径与基础设置 ---
# 尝试智能查找路径：优先使用代码目录，否则使用硬编码目录
CURRENT_DIR = Path(__file__).parent if '__file__' in locals() else Path.cwd()
DEFAULT_DICT_PATH = Path(r"E:\天之逸2025\Moze4.0\Moze Dict.xlsx")

if DEFAULT_DICT_PATH.exists():
    RULE_BOOK_PATH = DEFAULT_DICT_PATH
else:
    RULE_BOOK_PATH = CURRENT_DIR / "Moze Dict.xlsx"

TARGET_DIR = Path(r"E:\天之逸2025\Moze4.0\Moze4.0_Import")

CONFIG = {
    'TRANSFER_TARGET_1': '肖恩',
    'TRANSFER_TARGET_2': '工商银行(9579)',
    'CARD_PINGAN': '平安银行信用卡(4946)',
    'CARD_ICBC': '工商银行储蓄卡(9579)',
    'CARD_CMB': '招商银行储蓄卡(3379)',
    'CARD_CGB': '广发银行信用卡(4826)',
    'KEYWORD_CHARGING': '自助服务-充电桩',
    'KEYWORD_ALIPAY_CSV': '支付宝支付科技有限公司',
    'KEYWORD_WECHAT_XLSX': '微信支付账单明细',
}

# --- [配置 2] 关键词清单 (仅用于自动分类) ---
DATA_SOURCE = {
    'FRUIT': [
        "水果", "果园", "果蔬", "百果园", "鲜果", "果业", "苹果", "香蕉", "红枣", "大枣", "桃子", "水蜜桃", 
        "樱桃", "黄桃", "香梨", "雪梨", "柑橘", "沃柑", "草莓", "甘蔗", "桔子", "沙糖桔", "葡萄", "哈密瓜", 
        "西瓜", "甜瓜", "柚子", "柠檬", "橙子", "菠萝", "凤梨", "榴莲", "山竹", "蓝莓", "荔枝", "龙眼", 
        "椰子", "柿子", "杨梅", "李子", "芒果", "猕猴桃", "火龙果", "百香果", "莲雾", "车厘子", "牛油果",
        "鳄梨", "芭乐", "番石榴", "菠萝蜜", "桑葚", "枇杷", "杨桃", "无花果", "圣女果", "小番茄", "姑娘果",  
        "西梅", "青枣", "冬枣", "黑布林", "人参果", "丑八怪", "耙耙柑"
    ],
    'DRINK': ["饮料", "可乐", "红牛", "奶茶", "东鹏", "果汁", "椰汁", "酸奶", "咖啡", "拿铁", "乐虎", "AD钙奶", "甜酒汁"],
    'WATER': ["纯净水", "矿泉水", "农夫山泉", "怡宝", "百岁山", "娃哈哈", "今麦郎"],
    'SOFTWARE': ["软件", "APP", "应用"],
    'CHARGING': ["特来电", "星星充电", "小桔充电", "国家电网", "电费", "充电", "e充电", "云快充", "蔚来", "特斯拉", "超充"],
    
    'VEGETABLE': [
        "蔬菜",
        # 叶菜类
        "白菜", "菠菜","上海青", "油麦菜", "生菜", "娃娃菜", "空心菜", "苋菜", "菜芯", "韭菜", "香菜", "芹菜", "荠菜", "芥蓝",
         # 根茎类
        "土豆", "胡萝卜", "白萝卜", "青萝卜", "红薯", "紫薯", "山药", "芋头", "莲藕", "洋葱", "大蒜", "生姜", "竹笋", "芦笋", "茭白",
         # 茄果/瓜果类
        "西红柿", "茄子", "黄瓜", "西葫芦", "南瓜", "冬瓜", "苦瓜", "丝瓜", "青椒", "彩椒", "尖椒", "秋葵", "玉米",
        # 豆类
        "四季豆", "豇豆", "扁豆", "荷兰豆", "毛豆", "豌豆",  "蚕豆", "刀豆",
        # 菌菇类
        "香菇", "金针菇", "平菇", "杏鲍菇", "口蘑", "木耳", "银耳", "茶树菇", "猴头菇",
        # 葱姜蒜/调味菜
        "大葱", "小葱", "蒜苗", "蒜苔"  
    ],

    'Snack': [
        "零食", "切糕", "蛋糕", "面包", "腰果"
    ],
    'BEAN_PRODUCT': [
        "豆腐", "豆皮", "腐竹", "豆干", "豆腐泡", "千张", "豆卷", "腐皮", "油豆皮", "内脂豆腐", "老豆腐", 
        "嫩豆腐", "冻豆腐", "响铃卷", "素鸡"
    ],
    'INGREDIENTS': [
        "馒头", "火腿","广式腊肠", "榨菜",  "老干妈", "杂酱", "生水饺", "酱料", "炸酱",  "白砂糖", "蒸肉粉", "火锅底料",
        "食用盐", "生抽", "老抽",  "耗油", "料酒", "胡椒粉", "辣椒粉","辣椒酱",
        "鲜面条", "干面条", "挂面", "红豆", "绿豆", "黄豆", "花生米"
    ],
    'PORK': [
        "猪肉", "扇子骨", "板油", "五花肉", "梅花肉", "瘦肉", "猪蹄膀", "蹄膀", "排骨", "脊骨", "五花", 
        "前蹄", "大肠", "筒子骨", "棒骨", "猪油", "猪蹄", "荤油"
    ],
    'POULTRY': ["三黄鸡", "鸡腿", "s鸡翅", "鸡胸肉", "老母鸡", "乌鸡", "鸭肉", "鸭腿", "鸭翅", "鸭架", "老鸭"],
    'BEEF_MUTTON': ["牛肉", "牛腩", "牛排", "牛柳", "牛腱", "肥牛", "羊肉", "羊排", "羊腿"],
    'SEAFOOD': [
        "鱼", "虾", "蟹", "贝", "带鱼", "黄鱼", "鲈鱼", "鲫鱼", "草鱼", "基围虾", "皮皮虾", "大闸蟹", 
        "生蚝", "鱿鱼", "章鱼", "海带", "紫菜"
    ],
    'Eggs': ["鸡蛋", "土鸡蛋", "生咸鸭蛋", "皮蛋"],
    'COOKED': [
        "熟食", "水煮花生", "毛豆", "藕夹", "茄盒", "锅包肉", "猪头肉", "卤菜", "凉菜", "烧鸡", "烤鸭", 
        "酱牛肉", "熏鱼", "炸带鱼", "炸鱿鱼", "肉丸", "墨鱼丸", "牛肉丸", "鱼丸", "虾滑", "贡丸", "鱿鱼圈", 
        "糍粑", "熟咸鸭蛋", "熟肉肠", "糯米制品"
    ],
    'RICE': ["大米", "五常"],
    'MEAL': [
        "三镇民生", "兰州", "丝路", "沙县", "永和四喜", "水煎包", "老乡鸡", "混沌", "馄饨", "牛肉汤", "热干面", 
        "黄蜀郎", "鸡公煲", "小吃", "烧烤", "食堂", "路边摊", "麦香园", "长沙臭豆腐", "东苑一层", "东苑二层", 
        "西区食堂", "包子", "小笼包", "外勤", "板面", "水饺", "猪脚饭", "肠粉", "卤肉饭", "油泼面", "重庆小面", 
        "牛杂粉", "麻辣香锅", "麻辣烫", "煎包", "煎饼", "烧饼", "锅盔", "馕", "鸡蛋饼", "热干面", "牛肉面", 
        "刀削面", "盖浇饭", "炒饭", "汽水包", "煎豆折", "快餐", "小碗菜", "蒸饺", "餐厅"
    ],
    'Parking_fee': ["WF7023"],
    'DAILY_NECESSITIES': [
        "日常用品", 
        # 强特征词 (长词) - 优先覆盖
        "榨汁机", "挤压器", "水果压", "调料罐", "调料盒", "油罐", "盐罐", "料理机", "计时器", "切蛋器", "瓶起子",
        "洗发水", "沐浴露", "鼻毛修剪器", "掏耳勺", "清凉油", "花露水", "卫生纸", "洗洁精", "钢丝球", "洗衣粉", 
        "洁厕灵", "除锈剂", "清洗剂", "螺丝刀", "润滑油", "除锈剂", "理线器", "置物架", "凳子垫", "电磁炉",
        # 弱特征词 (单字) - 小心覆盖
        "锅", "勺", "刀",  "钩", "砧板", "擀面杖", "和面棒", "打蛋器", "压汁器", "筷", "碗", "杯", 
        "牙刷", "牙膏", "漱口杯", "香皂", "抽纸", "湿巾", "拖把", "扫把", "水桶", "盆", "围裙", 
        "卷尺", "胶水", "锁", "灯", "电池", "插座", "袋", "瓶", "盒", "壶", "罐", "桶", "纸", 
        "手套", "膜", "布", "衣架", "晒衣杆", "夹子", "挂钩", "雨伞"
    ],
    'Clothing_Shoes_Bags': ["袜子", "内裤", "帽子", "手套", "鞋", "T恤", "裤", "外套", "修裤脚"],
    'Adult_Products' : ["避孕套", "成人润滑剂", "安全套"],      
    'SERVER': ["节点", "Dler", "Dogess"],
    'Furniture_HomeTextiles': ["被子", "空调被", "枕头", "浴巾", "床笠"]
    
}

# --- [配置 3] 食材分类优先级 ---
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
    ('Clothing_Shoes_Bags', '', '服饰鞋包'),
    ('Furniture_HomeTextiles', '', '家具家纺'),
    ('SERVER', '节点', '虚拟其他') ,
    ('Adult_Products', 'Condoms', '保健用品')
]

# --- [配置 4] 全自动推导配置 ---
RAW_MAPPING_CONFIG = {
    ('饮食', '食'): ['午餐', '夜宵', '早餐', '晚餐', '纯净水', '零食', '食材', '饮料水果'],
    ('购物', '日用&家用'): ['家具家纺', '数码电器', '日常用品', '服饰鞋包'],
    ('交通', '通信&交通'): ['公共交通', '共享交通', '出租车', '汽车', '火车'],
    ('交通', '兼职'): ['加油充电'],
    ('居家', '日用&家用'): ['快递邮政', '物业费', '理发'],
    ('居家', '通信&交通'): ['电话费'],
    ('居家', '住'): ['房租', '水费', '电费'],
    ('居家', '食'): ['液化气费'],
    ('医疗', '日用&家用'): ['体检', '药品', '门诊'],
    ('医疗', 'Hormones'): ['保健用品'],
    ('娱乐', 'Hormones'): ['休闲保健', '住宿'],
    ('娱乐', '娱乐'): ['旅游度假', '电影', '网游电玩'],
    ('学习', '学习'): ['图书', '证书'],
    ('个人', 'Hormones'): ['The Girls'],
    ('个人', '工作'): ['个人其他', '保险'],
    ('个人', '兼职'): ['生意'],
    ('个人', '额外非必要开销'): ['孝敬', '礼金红包', '社交人情', '给予'],
    ('虚拟', '娱乐'): ['App', '虚拟其他', '订阅'],
    ('虚拟', '学习'): ['Software'],
    ('收入', '兼职'): ['二手折旧', '外卖跑腿(CNY)'],
    ('收入', '工作'): ['收入其他', '福利补贴', '薪资'],
    ('收入', '理财'): ['利息'],
    ('收入', ''): ['红包'],
    ('应收款项', ''): ['报账', '借出', '代付', '押金'],
    ('应付款项', ''): ['借入']
}

# --- 自动构建正则 ---
PATTERNS = {}
for key, words in DATA_SOURCE.items():
    PATTERNS[key] = r"(?:" + "|".join(map(re.escape, words)) + ")"

AUTO_MAP_DICT = {}
for (main_cat, project), sub_list in RAW_MAPPING_CONFIG.items():
    for sub_cat in sub_list:
        AUTO_MAP_DICT[sub_cat] = (main_cat, project)

COLUMN_MAPPING = {
    '交易时间': '交易时间', '交易类型': '交易类型', '交易对方': '交易对方', '商品': '商品',
    '收/支': '收/支', '金额(元)': '金额(元)', '金额（元）': '金额(元)',
    '支付方式': '支付方式', '当前状态': '当前状态', '备注': '备注',
    '交易单号': '交易单号', '商户单号': '商户单号',
    '交易分类': '交易类型', '商品说明': '商品', '金额': '金额(元)',
    '收/付款方式': '支付方式', '交易状态': '当前状态', '交易订单号': '交易单号',
    '商家订单号': '商户单号',
    '付款时间': '交易时间', '商品名称': '商品', '类型': '交易类型',
}

# ==========================================
#               核心逻辑函数
# ==========================================

def robust_date_converter(x):
    if isinstance(x, (datetime.datetime, datetime.date)): return pd.to_datetime(x)
    elif isinstance(x, (int, float)):
        try: return pd.to_datetime(x, unit='D', origin='1899-12-30')
        except: return pd.NaT
    elif isinstance(x, str): return pd.to_datetime(x, errors='coerce')
    return pd.NaT

def load_settings(rule_path: Path):
    global CONFIG
    if not rule_path.exists(): return
    try:
        df = pd.read_excel(rule_path, sheet_name='Settings', engine='openpyxl')
        df.columns = df.columns.str.strip().str.lower()
        k_col = next((c for c in df.columns if 'key' in c or '键' in c), None)
        v_col = next((c for c in df.columns if 'value' in c or '值' in c), None)
        if k_col and v_col:
            for _, row in df.iterrows():
                k, v = str(row[k_col]).strip(), str(row[v_col]).strip()
                if k in CONFIG and v and v.lower() != 'nan': CONFIG[k] = v
            print(f"已加载自定义配置")
    except: pass

def load_rules(rule_path: Path):
    print("正在加载商家映射规则...", flush=True)
    try:
        xl = pd.ExcelFile(rule_path, engine='openpyxl')
        sheet_name = None
        for s in xl.sheet_names:
            if s.lower() == 'settings': continue
            tmp = pd.read_excel(rule_path, sheet_name=s, nrows=5, engine='openpyxl')
            if '商家(old)' in tmp.columns: sheet_name = s; break
        
        if not sheet_name: print("错：没找到商家规则表"); return None
        df = pd.read_excel(rule_path, sheet_name=sheet_name, engine='openpyxl')
        if 'is_regex' not in df.columns: df['is_regex'] = 0
        df['is_regex'] = pd.to_numeric(df['is_regex'], errors='coerce').fillna(0)
        df['商家(old)'] = df['商家(old)'].astype(str).str.strip()
        return df
    except Exception as e: print(f"加载规则失败: {e}"); return None

def sniff_and_load_data(file_path: Path):
    print(f"读取: {file_path.name}", flush=True)
    ftype, skip, enc, df = None, 0, None, None
    
    if file_path.suffix.lower() == '.csv':
        for e in ['gbk', 'gb18030', 'utf-8-sig', 'utf-8']:
            try:
                with open(file_path, 'r', encoding=e, errors='strict') as f:
                    for i in range(50):
                        line = f.readline()
                        if CONFIG['KEYWORD_ALIPAY_CSV'] in line:
                            ftype, skip, enc = "AliPay", i+1, e
                            break
                if ftype: break
            except UnicodeDecodeError: continue
            except Exception: break
            
    elif file_path.suffix.lower() == '.xlsx':
        try:
            tmp = pd.read_excel(file_path, header=None, nrows=20, engine='openpyxl')
            if CONFIG['KEYWORD_WECHAT_XLSX'] in str(tmp.iloc[0,0]): ftype, skip = "WeChat", 16
        except: pass

    if not ftype: 
        print("  跳过：无法识别的文件类型。")
        return None

    try:
        if ftype == "WeChat": df = pd.read_excel(file_path, skiprows=skip, engine='openpyxl')
        else: df = pd.read_csv(file_path, skiprows=skip, encoding=enc)
        df.columns = df.columns.str.strip()
        df.rename(columns=COLUMN_MAPPING, inplace=True)
        if '金额(元)' not in df.columns: df['金额(元)'] = 0
        df['金额(元)'] = pd.to_numeric(df['金额(元)'].astype(str).str.replace('¥','').str.strip(), errors='coerce').fillna(0)
        df['_source_tag'] = '#WechatPay' if ftype == "WeChat" else '#AliPay'
        print(f"  成功读取 {len(df)} 行。")
        return df
    except Exception as e: print(f"  加载错: {e}"); return None

def ensure_columns(df, main_col, sub_col):
    required_cols = [main_col, sub_col, '金额', '记录类型', '账户', '商家', '描述', '项目', '名称', '对象', '标签', '日期', '时间', '币种', '手续费', '折扣']
    for c in required_cols:
        if c not in df.columns: df[c] = pd.NA
    str_cols = ['记录类型', '项目', '对象', main_col, sub_col, '商家', '名称']
    for c in str_cols:
        df[c] = df[c].fillna("")
    return df

def _create_transfer_pair(df_subset, out_acc, in_acc):
    if df_subset.empty: return []
    out_df = df_subset.copy()
    out_df['金额'] = pd.to_numeric(out_df['金额(元)']) * -1
    out_df['记录类型'] = '转出'; out_df['账户'] = out_acc
    in_df = df_subset.copy()
    in_df['金额'] = pd.to_numeric(in_df['金额(元)']); in_df['记录类型'] = '转入'; in_df['账户'] = in_acc
    return [out_df, in_df]

def process_transfers(df, main_col, sub_col):
    res = []
    t1 = CONFIG['TRANSFER_TARGET_1']
    t2 = CONFIG['TRANSFER_TARGET_2']
    
    res.extend(_create_transfer_pair(df[(df["交易对方"] == t1) & (df["收/支"] == "收入")], '零钱2', '零钱3'))
    res.extend(_create_transfer_pair(df[(df["交易对方"] == t1) & (df["收/支"] == "支出")], '零钱3', '零钱2'))

    icbc_pairs = _create_transfer_pair(df[df["交易对方"] == t2], '零钱3', '工商银行')
    for p in icbc_pairs: p[sub_col] = '提现'
    res.extend(icbc_pairs)

    if not res: return pd.DataFrame()
    
    ret = pd.concat(res)
    ret = ensure_columns(ret, main_col, sub_col)
    ret[main_col] = "转账"
    ret[sub_col] = ret[sub_col].replace("", "转账")
    
    ret['日期'] = ret['交易时间'].dt.strftime('%Y/%m/%d')
    ret['时间'] = ret['交易时间'].dt.strftime('%H:%M:%S')
    ret['_Sort_Date'] = ret['交易时间']
    ret['币种'] = "CNY"; ret['手续费'] = 0; ret['折扣'] = 0
    
    return ret

def process_heuristics(df, main_col, sub_col):
    df = ensure_columns(df, main_col, sub_col)
    
    for col in [sub_col, '项目', '描述', '商家', '商品']:
        df[col] = df[col].astype(str).str.strip().replace("nan", "")
    
    def check_keywords(key_name):
        if key_name not in PATTERNS: return pd.Series([False] * len(df))
        pattern = PATTERNS[key_name]
        mask = (
            df['商家'].str.contains(pattern, regex=True, na=False) |
            df['商品'].str.contains(pattern, regex=True, na=False) |
            df['描述'].str.contains(pattern, regex=True, na=False)
        )
        if '备注' in df.columns:
            mask |= df['备注'].astype(str).str.contains(pattern, regex=True, na=False)
        return mask

    uncat = (df[sub_col] == "")
    allow_overwrite = (df[sub_col] == "") | (df[sub_col] == "食材")

    # 1. 充电桩
    mask = df['商品'].str.contains(CONFIG['KEYWORD_CHARGING'], na=False)
    mask_cat = mask & uncat
    if mask_cat.any(): 
        df.loc[mask_cat, ['名称', main_col, sub_col]] = ['充电', '交通', '加油充电']

    # 2. 软件
    mask = check_keywords('SOFTWARE') & uncat
    if mask.any():
        df.loc[mask, sub_col] = 'Software'
    
    # 3. 水果
    mask = check_keywords('FRUIT') & (~mask) & allow_overwrite
    if mask.any(): df.loc[mask, ['名称', sub_col]] = ['水果', '饮料水果']
    
    # 4. 饮料
    mask = check_keywords('DRINK') & (~mask) & allow_overwrite
    if mask.any():
        df.loc[mask, ['名称', sub_col]] = ['饮料', '饮料水果']
    
    # 5. 纯净水
    mask = check_keywords('WATER') & (~mask) & allow_overwrite
    if mask.any(): 
        df.loc[mask, sub_col] = '纯净水'
        df.loc[mask, '名称'] = ""

    # 6. 大类识别
    mask_is_meal = check_keywords('MEAL')
    mask_is_cooked = check_keywords('COOKED')
    
    for key, name, sub_c in INGREDIENT_PRIORITY:
        mask = check_keywords(key)
        if key in ['SEAFOOD', 'PORK', 'POULTRY', 'BEEF_MUTTON', 'VEGETABLE']:
            mask = mask & (~mask_is_cooked)

        mask = mask & (~mask_is_meal) & (uncat | (df[main_col].isin(['购物', '居家', '饮食'])))
        
        if mask.any():
            # [修正] 之前这里误用了 sub_col (表头名)
            df.loc[mask, sub_col] = sub_c  
            df.loc[mask, '名称'] = name
            df[sub_col] = df[sub_col].replace("nan", "")

    # 7. 日常用品识别 (强弱双轨)
    daily_keywords = DATA_SOURCE['DAILY_NECESSITIES']
    strong_keywords = [w for w in daily_keywords if len(w) > 1]
    weak_keywords = [w for w in daily_keywords if len(w) == 1]
    
    pattern_strong = r"(?:" + "|".join(map(re.escape, strong_keywords)) + ")"
    pattern_weak = r"(?:" + "|".join(map(re.escape, weak_keywords)) + ")"
    
    def get_mask(pat):
        m = (
            df['商家'].str.contains(pat, regex=True, na=False) |
            df['商品'].str.contains(pat, regex=True, na=False) |
            df['描述'].str.contains(pat, regex=True, na=False)
        )
        if '备注' in df.columns:
            m |= df['备注'].astype(str).str.contains(pat, regex=True, na=False)
        return m

    # 7.1 强特征 (优先覆盖)
    mask_strong = get_mask(pattern_strong)
    mask_strong = mask_strong & (~mask_is_meal)
    if mask_strong.any():
        df.loc[mask_strong, [main_col, sub_col]] = ['购物', '日常用品']
        df.loc[mask_strong, '名称'] = ''
        df[sub_col] = df[sub_col].replace("nan", "")

    # 7.2 弱特征 (保护食材/餐饮)
    mask_weak = get_mask(pattern_weak)
    mask_weak = mask_weak & (df[sub_col] != '日常用品')
    mask_weak = mask_weak & (uncat | (df[main_col].isin(['购物', '居家', '饮食'])))
    mask_weak = mask_weak & (~df[sub_col].isin(['食材', '饮料水果'])) 
    mask_weak = mask_weak & (~mask_is_meal)

    if mask_weak.any():
        df.loc[mask_weak, [main_col, sub_col]] = ['购物', '日常用品']
        df.loc[mask_weak, '名称'] = ''
        df[sub_col] = df[sub_col].replace("nan", "")
        
    # [优化] 停车费报销 (针对 WF7023 - 显式修正，防止被后续逻辑误覆盖)
    mask = check_keywords('Parking_fee')
    if mask.any():
        df.loc[mask, '记录类型'] = "应收款项"
        df.loc[mask, main_col] = '应收款项'
        df.loc[mask, sub_col] = '报账'
        df.loc[mask, '名称'] = "停车费"
        df.loc[mask, '对象'] ="天之逸"
        df.loc[mask, '项目'] =""

    # 8. 全自动推导 (根据子类推主类和项目)
    mask_main = (df[main_col] == "")
    df.loc[mask_main, main_col] = df.loc[mask_main, sub_col].map(lambda x: AUTO_MAP_DICT.get(x, (None, None))[0]).fillna("")

    mask_proj = (df['项目'] == "")
    df.loc[mask_proj, '项目'] = df.loc[mask_proj, sub_col].map(lambda x: AUTO_MAP_DICT.get(x, (None, None))[1]).fillna("")

    # ==========================================
    # 9. [新增/优化] 智能债权债务识别 & 强制兜底
    # ==========================================
    
    # 组1：应收 (Receivables)
    keys_rec = ['报账', '借出', '代付', '押金']
    # 组2：应付 (Payables)
    keys_pay = ['借入']
    
    # 合并关键词用于正则
    all_debt_keys = keys_rec + keys_pay
    
    # A. 智能描述提取 (提取 "动词+人名"，如 "代付黄玲")
    debt_pat = rf"({'|'.join(all_debt_keys)})\s*(.*)"
    extracted = df['描述'].str.extract(debt_pat, expand=True)
    mask_found = extracted[0].notna()

    if mask_found.any():
        # 1. 填入子类别
        df.loc[mask_found, sub_col] = extracted[0]
        
        # 2. 智能填入对象
        extracted_names = extracted[1].str.strip()
        mask_with_name = mask_found & (extracted_names != "")
        if mask_with_name.any():
            df.loc[mask_with_name, '对象'] = extracted_names[mask_with_name]
            
        # 3. 同时清空项目和描述 (因为信息已提取，不再需要)
        df.loc[mask_found, ['项目', '描述']] = ""

    # B. 强制兜底逻辑 (根据子类别修正记录类型)
    
    # 处理【应收款项】
    mask_is_rec = df[sub_col].isin(keys_rec)
    if mask_is_rec.any():
        df.loc[mask_is_rec, ['记录类型', main_col]] = ['应收款项', '应收款项']

    # 处理【应付款项】
    mask_is_pay = df[sub_col].isin(keys_pay)
    if mask_is_pay.any():
        df.loc[mask_is_pay, ['记录类型', main_col]] = ['应付款项', '应付款项']

    return df

def process_main(df, df_rules, main_col, sub_col):
    excl = ["已全额退款", "交易关闭"]
    mask = (~df.get("当前状态","").isin(excl)) & (df.get("收/支") == "支出")
    df = df[mask].copy()
    if df.empty: return pd.DataFrame()

    cnt_wx = (df['_source_tag'] == '#WechatPay').sum()
    cnt_ali = (df['_source_tag'] == '#AliPay').sum()
    print(f"--- 处理主交易 ({len(df)}条) [微信: {cnt_wx} | 支付宝: {cnt_ali}] ---", flush=True)

    df = ensure_columns(df, main_col, sub_col)

    # 1. 提取正则模式（注意外面的括号）
    (PINGAN_pat, ICBC_pat, CMB_pat, CGB_pat) = (
        re.escape(CONFIG['CARD_PINGAN']),
        re.escape(CONFIG['CARD_ICBC']),
        re.escape(CONFIG['CARD_CMB']),
        re.escape(CONFIG['CARD_CGB'])
    )

# 2. 执行替换
    df['支付方式'] = df['支付方式'].astype(str).replace({
        rf'^(?:{PINGAN_pat}).*': '平安银行4946',
        rf'^(?:{ICBC_pat}).*':   '工商银行',
        rf'^(?:{CMB_pat}).*':    '招商银行Ⅱ',
        rf'^(?:{CGB_pat}).*':    '广发银行4826',
        r'^零钱.*':              '零钱3'
    }, regex=True)

    df['商家(old)'] = df.get('交易对方', pd.NA).astype(str).str.strip()
    rename_map = {c: f'{c}_rule' for c in df_rules.columns if c not in ['商家(old)','is_regex']}
    exact = df_rules[df_rules['is_regex']==0]
    df = pd.merge(df, exact.rename(columns=rename_map).drop(columns='is_regex', errors='ignore'), on='商家(old)', how='left')
    
    regex = df_rules[df_rules['is_regex']==1]
    rule_col = f"{main_col}_rule"
    uncat = df[df[rule_col].isna()].index if rule_col in df.columns else df.index
    
    if not uncat.empty and not regex.empty:
        to_match = df.loc[uncat, '交易对方'].astype(str)
        for _, r in regex.iterrows():
            if to_match.empty: break
            try:
                m = to_match.str.contains(r['商家(old)'], regex=True)
                idx = to_match[m].index
                if not idx.empty:
                    for c in rename_map.keys(): df.loc[idx, f'{c}_rule'] = r[c]
                    to_match = to_match.drop(idx)
            except: pass

    # --- 1. 合并规则 (暂不合并描述) ---
    for orig, rule in rename_map.items():
        if rule in df.columns:
            if orig == '描述': continue 
            df[orig] = df[rule].combine_first(df[orig])
            df.drop(columns=[rule], inplace=True)

    # --- 2. 保留淘宝原有逻辑 (设置商家) ---
    if '商户单号' in df.columns:
        mask_tb = (df['_source_tag']=='#AliPay') & df['商户单号'].astype(str).str.startswith('T200P4', na=False)
        if mask_tb.any(): df.loc[mask_tb, '商家'] = '淘宝'
            
    # ==========================================
    # [核心逻辑] 描述：备注 > 字典 > (淘宝?商品名:空)
    # ==========================================
    
    # 基础值：空值
    df['描述'] = ""
    
    # 步骤 A：如果是淘宝，预填充商品名
    mask_is_taobao = (df['商家'] == '淘宝')
    if mask_is_taobao.any():
        df.loc[mask_is_taobao, '描述'] = df.loc[mask_is_taobao, '商品'].fillna("").astype(str)
    
    # 步骤 B：如果有备注，优先使用备注
    if '备注' in df.columns:
        raw_memo = df['备注'].astype(str).replace('nan', '').str.strip()
        df['描述'] = np.where(raw_memo != "", raw_memo, df['描述'])

    # 步骤 C：如果有字典描述规则，强制覆盖
    if '描述_rule' in df.columns:
        df['描述'] = df['描述_rule'].combine_first(df['描述'])
        df.drop(columns=['描述_rule'], inplace=True, errors='ignore')

    # 步骤 D：清洗 "/" 为空值
    df.loc[df['描述'].astype(str).str.strip() == '/', '描述'] = ""

    # [微调] 仅去掉首尾空格和换行符
    df['描述'] = df['描述'].astype(str).str.replace(r'[\r\n]+', ' ', regex=True).str.strip()

    # ==========================================

    df['记录类型'] = df.get('记录类型', pd.NA).replace("", pd.NA).fillna(df['收/支'])
    df['商家'] = df.get('商家', pd.NA).replace("", pd.NA).fillna(df['交易对方'])

    df = ensure_columns(df, main_col, sub_col)

    # 餐饮时段识别
    df[sub_col] = df[sub_col].astype(str).str.strip().replace("nan", "")
    uncat_mask = (df[sub_col] == "")
    
    is_meal = (
        df['商家(old)'].str.contains(PATTERNS['MEAL'], regex=True, na=False) | 
        df['商品'].str.contains(PATTERNS['MEAL'], regex=True, na=False) |
        df['商家'].str.contains(PATTERNS['MEAL'], regex=True, na=False) |
        df['描述'].str.contains(PATTERNS['MEAL'], regex=True, na=False)
    ) & uncat_mask
    
    if is_meal.any():
        h = df.loc[is_meal, '交易时间'].dt.hour
        vals = np.select([(h>=6)&(h<11), (h>=11)&(h<16), (h>=16)&(h<21)], ["早餐", "午餐", "晚餐"], default="夜宵")
        df.loc[is_meal, sub_col] = vals

    df = process_heuristics(df, main_col, sub_col)
    
    raw_amt = df.get('金额(元)', 0)
    df['金额'] = pd.to_numeric(raw_amt, errors='coerce').fillna(0)
    
    # 默认为负
    neg_mask = df['记录类型'].isin(['支出','应付款项','应收款项']) & (df.get('收/支')=='支出')
    df.loc[neg_mask, '金额'] *= -1
    
    # [新增] 强制修正：借入金额必须为正
    mask_borrow = (df[sub_col] == '借入')
    if mask_borrow.any():
        df.loc[mask_borrow, '金额'] = df.loc[mask_borrow, '金额'].abs()
    
    fix_mask = (df['支付方式'] == "/") & (df['记录类型'] == "收入")
    df.loc[fix_mask, '支付方式'] = "零钱3"
    
    df['日期'] = df['交易时间'].dt.strftime('%Y/%m/%d')
    df['时间'] = df['交易时间'].dt.strftime('%H:%M:%S')
    df['_Sort_Date'] = df['交易时间']
    df['账户'] = df.get('支付方式', ""); df['币种'] = "CNY"; df['手续费'] = 0; df['折扣'] = 0
    
    rp_mask = df['记录类型'].isin(['应收款项', '应付款项'])
    df.loc[rp_mask, '商家'] = ""
    
    return df

def main():
    print(f"{bcolors.BOLD}=== Moze 导入工具 v11.12 (智能债权修复版) ==={bcolors.ENDC}")
    try:
        load_settings(RULE_BOOK_PATH)
        df_rules = load_rules(RULE_BOOK_PATH)
        if df_rules is None: input("按 Enter 退出"); return

        root = tk.Tk(); root.withdraw()
        print("\n选择文件...")
        files = filedialog.askopenfilenames(filetypes=[("Excel/CSV", "*.xlsx;*.csv")])
        if not files: return
        
        start = simpledialog.askstring("筛选", "起始日期 yymmdd (留空导入全部):")
        st_date = None
        if start: 
            try: st_date = datetime.datetime.strptime(start.strip(), "%y%m%d")
            except: pass

        cols = [c for c in df_rules.columns if c not in ['商家(old)','is_regex']]
        main_col = next((c for c in cols if "主类" in c), None)
        sub_col = next((c for c in cols if "子类" in c), None)
        if not main_col: print("错：找不到主类列"); return

        dfs = [d for f in files if (d := sniff_and_load_data(Path(f))) is not None]
        if not dfs: return
        df_raw = pd.concat(dfs, ignore_index=True)
        
        if '交易时间' not in df_raw.columns: df_raw['交易时间'] = pd.NaT
        df_raw['交易时间'] = df_raw['交易时间'].apply(robust_date_converter)
        if st_date: df_raw = df_raw[df_raw['交易时间'] >= st_date]
        df_raw['交易对方'] = df_raw.get('交易对方', pd.NA).astype(str).str.strip()

        targets = [CONFIG['TRANSFER_TARGET_1'], CONFIG['TRANSFER_TARGET_2']]
        mask_trans = df_raw['交易对方'].isin(targets)
        
        res_dfs = []
        if mask_trans.any(): res_dfs.append(process_transfers(df_raw[mask_trans].copy(), main_col, sub_col))
        if (~mask_trans).any(): res_dfs.append(process_main(df_raw[~mask_trans].copy(), df_rules, main_col, sub_col))
        
        if not res_dfs: print("无结果"); return
        df_final = pd.concat(res_dfs, ignore_index=True)
        
        for c in cols: 
            if c not in df_final.columns: df_final[c] = ""
        df_final = df_final.fillna("")
            
        if '标签' in df_final.columns and '_source_tag' in df_final.columns:
            m = ~df_final['记录类型'].isin(['转入', '转出'])
            df_final.loc[m, '标签'] = (df_final.loc[m, '标签'] + " " + df_final.loc[m, '_source_tag']).str.strip()

        df_final['_T_Sort'] = df_final['记录类型'].map({'转出':0, '转入':1}).fillna(9)
        df_final.sort_values(['_Sort_Date', '_T_Sort'], ascending=[False, True], inplace=True)

        if not TARGET_DIR.exists(): TARGET_DIR.mkdir(parents=True, exist_ok=True)
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = TARGET_DIR / f"MOZE导入_{ts}.csv"
        df_final.to_csv(path, index=False, columns=cols, encoding='utf-8-sig')
        
        print(f"\n{bcolors.OKGREEN}成功! 文件: {path}{bcolors.ENDC}")
        
        # --- 数据完整性检查 ---
        print(f"\n{bcolors.BOLD}--- 数据完整性检查 ---{bcolors.ENDC}")
        
        mask_exp = df_final['记录类型'] == '支出'
        invalid_exp = mask_exp & (
            (df_final[main_col] == "") | (df_final[main_col].isna()) |
            (df_final[sub_col] == "") | (df_final[sub_col].isna()) |
            (df_final['项目'] == "") | (df_final['项目'].isna())
        )
        
        if invalid_exp.any():
            print(f"{bcolors.FAIL}[严重] {invalid_exp.sum()} 条【支出】记录缺失分类或项目:{bcolors.ENDC}")
            print(df_final.loc[invalid_exp, ['日期', '商家', '描述', '金额', sub_col, '项目']].head(5).to_string())
        else:
            print(f"{bcolors.OKGREEN}√ 支出记录完美{bcolors.ENDC}")

        mask_rec = df_final['记录类型'].isin(['应收款项', '应付款项'])
        invalid_rec = mask_rec & (
            (df_final[main_col] == "") | (df_final[main_col].isna()) |
            (df_final[sub_col] == "") | (df_final[sub_col].isna()) |
            (df_final['对象'] == "") | (df_final['对象'].isna())
        )
        
        if invalid_rec.any():
            print(f"{bcolors.FAIL}[严重] {invalid_rec.sum()} 条【应收/应付】记录缺失分类或对象:{bcolors.ENDC}")
            print(df_final.loc[invalid_rec, ['日期', '商家', '描述', '金额', '对象']].head(5).to_string())
        else:
            print(f"{bcolors.OKGREEN}√ 应收/应付记录完美{bcolors.ENDC}")

    except Exception: traceback.print_exc()

if __name__ == "__main__": main(); input("\n回车退出...")