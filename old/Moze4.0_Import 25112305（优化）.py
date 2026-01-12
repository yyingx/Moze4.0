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
#      Moze 导入脚本 v10.25 (店铺+备注拼接版)
# ==========================================

# --- [配置区 1] 终端颜色 ---
class bcolors:
    OKGREEN = '\033[92m'; WARNING = '\033[93m'; FAIL = '\033[91m'; ENDC = '\033[0m'; BOLD = '\033[1m'

# --- [配置区 2] 路径与关键词 ---
RULE_BOOK_PATH = Path(r"E:\天之逸2025\Moze4.0\Moze Dict.xlsx")
TARGET_DIR = Path(r"E:\天之逸2025\Moze4.0\Moze4.0_Import")

CONFIG = {
    'TRANSFER_TARGET_1': '肖恩',
    'TRANSFER_TARGET_2': '工商银行(9579)',
    'CARD_PINGAN': '平安银行信用卡(4946)',
    'CARD_ICBC': '工商银行储蓄卡(9579)',
    'KEYWORD_CHARGING': '自助服务-充电桩',
    'KEYWORD_ALIPAY_CSV': '支付宝支付科技有限公司',
    'KEYWORD_WECHAT_XLSX': '微信支付账单明细',
}

# --- [配置区 3] 正则表达式库 ---
PATTERNS = {
    'FRUIT': r"(?:水果|果园|果蔬|百果园|鲜果|果业|苹果|香蕉|红枣|大枣|桃子|水蜜桃|樱桃|黄桃|香梨|雪梨|柑橘|沃柑|草莓|甘蔗|桔子|沙糖桔|葡萄|哈密瓜|西瓜|甜瓜|柚子|柠檬|橙子|菠萝|凤梨|榴莲|山竹|蓝莓|荔枝|龙眼|椰子|柿子|杨梅|李子|芒果|猕猴桃|火龙果|百香果|莲雾|车厘子|牛油果|鳄梨|芭乐|番石榴|菠萝蜜|桑葚|枇杷|杨桃|无花果|圣女果|小番茄|姑娘果|西梅|青枣|冬枣|黑布林|人参果|释迦)",
    'DRINK': r"(?:可乐|红牛|奶茶|东鹏|饮料|汽水|果汁|椰汁|酸奶)",
    'WATER': r"(?:纯净水|矿泉水|农夫山泉|怡宝|百岁山|娃哈哈|今麦郎)",
    'SOFTWARE': r"(?:软件)",
    # 蔬菜大全
    'VEGETABLE': r"(?:青菜|白菜|娃娃菜|菠菜|生菜|油麦菜|韭菜|芹菜|香菜|土豆|茄子|豆角|辣椒|青椒|洋葱|西红柿|番茄|黄瓜|冬瓜|南瓜|苦瓜|丝瓜|萝卜|胡萝卜|藕|莲藕|笋|莴笋|蘑菇|香菇|金针菇|木耳|豆芽|玉米|红薯|山药|芋头)",
    # 豆制品
    'BEAN_PRODUCT': r"(?:豆腐|豆皮|腐竹|豆干|豆腐泡|素鸡|千张|豆卷|腐皮|油豆皮|内脂豆腐|老豆腐|嫩豆腐|冻豆腐|响铃卷)",
    # 食材与调料
    'INGREDIENTS': r"(?:馒头|火腿|葱|姜|蒜|酱料|料酒|白砂糖|生抽|老抽|蒸肉粉|榨菜)",
    # 熟食
    'COOKED': r"(?:水煮花生|毛豆|藕夹|茄盒|锅包肉|猪头肉)",
    # 大米
    'RICE': r"(?:大米)",
    # 正餐
    'MEAL': r"(?:三镇民生|兰州|丝路|沙县|永和四喜|水煎包|老乡鸡|混沌|馄饨|牛肉汤|热干面|黄蜀郎|鸡公煲|小吃|烧烤|食堂|路边摊|麦香园|长沙臭豆腐|东苑一层|东苑二层|西区食堂|包子|小笼包|外勤|板面|水饺|猪脚饭|肠粉|卤肉饭|油泼面|重庆小面|牛杂粉|麻辣香锅|麻辣烫|煎包|煎饼|烧饼|锅盔|馕|土豆鸡蛋饼)",
    # 垃圾字符库 (保留强力清洗，去掉 / 和 G代码)
    'CLEAN': r"(?:2\.00E\+15|经营码交易|余额充值|.*?后付|消费|费用|支付|商户|二维码收款|收款|美团|编号|转账|错误|订单编号|\d{10,}|/|[A-Za-z]\d+-\d+)"
}

# --- [配置区 4] 全自动推导引擎 ---
AUTO_MAP_DICT = {
    # --- 个人 ---
    'The Girls': ('个人', 'Hormones'), '个人其他': ('个人', '工作'), '保险': ('个人', '工作'),
    '孝敬': ('个人', '额外非必要开销'), '生意': ('个人', '兼职'), '礼金红包': ('个人', '额外非必要开销'),
    '社交人情': ('个人', '额外非必要开销'), '给予': ('个人', '额外非必要开销'),
    # --- 交通 ---
    '公共交通': ('交通', '通信&交通'), '共享交通': ('交通', '通信&交通'), '出租车': ('交通', '通信&交通'),
    '加油充电': ('交通', '兼职'), '汽车': ('交通', '通信&交通'), '火车': ('交通', '通信&交通'),
    # --- 医疗 ---
    '体检': ('医疗', '日用&家用'), '保健用品': ('医疗', 'Hormones'),
    '药品': ('医疗', '日用&家用'), '门诊': ('医疗', '日用&家用'),
    # --- 娱乐 ---
    '休闲保健': ('娱乐', 'Hormones'), '住宿': ('娱乐', 'Hormones'), '旅游度假': ('娱乐', '娱乐'),
    '电影': ('娱乐', '娱乐'), '网游电玩': ('娱乐', '娱乐'),
    # --- 学习 ---
    '图书': ('学习', '学习'), '证书': ('学习', '学习'),
    # --- 居家 ---
    '快递邮政': ('居家', '日用&家用'), '房租': ('居家', '住'), '水费': ('居家', '住'),
    '液化气费': ('居家', '食'), '物业费': ('居家', '日用&家用'), '理发': ('居家', '日用&家用'),
    '电话费': ('居家', '通信&交通'), '电费': ('居家', '住'),
    # --- 收入 ---
    '二手折旧': ('收入', '兼职'), '利息': ('收入', '理财'), '外卖跑腿(CNY)': ('收入', '兼职'),
    '收入其他': ('收入', '工作'), '福利补贴': ('收入', '工作'), '红包': ('收入', ''), '薪资': ('收入', '工作'),
    # --- 虚拟 ---
    'App': ('虚拟', '娱乐'), 'Software': ('虚拟', '学习'), '虚拟其他': ('虚拟', '娱乐'), '订阅': ('虚拟', '娱乐'),
    # --- 购物 ---
    '共享租赁': ('购物', '兼职'), '大件': ('购物', '兼职'), '家具家纺': ('购物', '日用&家用'),
    '摄影文印': ('购物', '兼职'), '数码电器': ('购物', '日用&家用'),
    '日常用品': ('购物', '日用&家用'), '服饰鞋包': ('购物', '日用&家用'),
    # --- 饮食 ---
    '午餐': ('饮食', '食'), '夜宵': ('饮食', '食'), '早餐': ('饮食', '食'), '晚餐': ('饮食', '食'),
    '纯净水': ('饮食', '食'), '零食': ('饮食', '食'), '食材': ('饮食', '食'), '饮料水果': ('饮食', '食'),
}

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
    print("正在加载商家规则...", flush=True)
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
    
    # 关键词搜索函数
    def check_keywords(pattern):
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
    if mask.any():
        split = df.loc[mask, '商品'].str.split('/', n=1, expand=True)
        if not split.empty and split.shape[1]>1: 
            df.loc[mask, '描述'] = split[1].str.strip()
        mask_cat = mask & uncat
        if mask_cat.any(): df.loc[mask_cat, ['名称', main_col, sub_col]] = ['充电', '交通', '加油充电']

    # 2. 软件
    mask = check_keywords(PATTERNS['SOFTWARE']) & uncat
    if mask.any(): df.loc[mask, sub_col] = 'Software'
    
    # 3. 水果
    mask = check_keywords(PATTERNS['FRUIT']) & (~mask) & allow_overwrite
    if mask.any(): df.loc[mask, ['名称', sub_col]] = ['水果', '饮料水果']
    
    # 4. 饮料
    mask = check_keywords(PATTERNS['DRINK']) & (~mask) & allow_overwrite
    if mask.any(): df.loc[mask, ['名称', sub_col]] = ['饮料', '饮料水果']
    
    # 5. 纯净水
    mask = check_keywords(PATTERNS['WATER']) & (~mask) & allow_overwrite
    if mask.any(): 
        df.loc[mask, sub_col] = '纯净水'
        df.loc[mask, '名称'] = ""

    # 6. 大类识别
    for key, name in [('RICE', '大米'), ('COOKED', '熟食'), ('BEAN_PRODUCT', '豆制品'), ('VEGETABLE', '蔬菜')]:
        mask = check_keywords(PATTERNS[key])
        mask = mask & (uncat | (df[main_col].isin(['购物', '居家'])))
        if mask.any():
            df.loc[mask, sub_col] = '食材'
            df.loc[mask, '名称'] = name
            df[sub_col] = df[sub_col].replace("nan", "")

    # 7. 食材与调料
    mask = check_keywords(PATTERNS['INGREDIENTS'])
    mask_already_set = df['名称'].isin(['大米', '熟食', '豆制品', '蔬菜'])
    mask = mask & (~mask_already_set) & (uncat | (df[main_col].isin(['购物', '居家'])))
    if mask.any():
        df.loc[mask, sub_col] = '食材'
        df.loc[mask, '名称'] = "" 
        df[sub_col] = df[sub_col].replace("nan", "")

    # 8. [全自动推导]
    mask_main = (df[main_col] == "")
    df.loc[mask_main, main_col] = df.loc[mask_main, sub_col].map(lambda x: AUTO_MAP_DICT.get(x, (None, None))[0]).fillna("")

    mask_proj = (df['项目'] == "")
    df.loc[mask_proj, '项目'] = df.loc[mask_proj, sub_col].map(lambda x: AUTO_MAP_DICT.get(x, (None, None))[1]).fillna("")

    # 9. 最终垃圾字符清洗 (不会删掉正常文字，只会删掉 / 或 Gxxxx)
    df['描述'] = df['描述'].astype(str).str.replace(PATTERNS['CLEAN'], ' ', regex=True)
    df['描述'] = df['描述'].str.strip()
    
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

    p_pat, i_pat = re.escape(CONFIG['CARD_PINGAN']), re.escape(CONFIG['CARD_ICBC'])
    df['支付方式'] = df['支付方式'].astype(str).replace({
        rf'^(?:{p_pat}).*': '平安银行4946', rf'^(?:{i_pat}).*': '工商银行', r'^零钱.*': '零钱3'
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

    for orig, rule in rename_map.items():
        if rule in df.columns:
            df[orig] = df[rule].combine_first(df[orig])
            df.drop(columns=[rule], inplace=True)

    if '商户单号' in df.columns:
        mask_tb = (df['_source_tag']=='#AliPay') & df['商户单号'].astype(str).str.startswith('T200P4', na=False)
        if mask_tb.any(): df.loc[mask_tb, '商家'] = '淘宝'
            
    # ==========================================
    # [核心拼接逻辑] 描述 = 商品(店铺) + "." + 备注
    # ==========================================
    
    # 1. 基础描述 = 商品名 (通常是店铺名)
    df['描述'] = df.get('描述', pd.NA).fillna(df.get('商品', pd.NA))
    
    if '备注' in df.columns:
        # 2. 拿到原始备注，先清洗掉垃圾字符 (Gxxx 或 /)
        # 目的：如果备注只有垃圾字符，清洗后为空，就不会被拼上去
        raw_memo = df['备注'].fillna("").astype(str).str.strip()
        cleaned_memo_check = raw_memo.str.replace(PATTERNS['CLEAN'], '', regex=True).str.strip()
        
        # 3. 找出有效备注 (清洗后不为空)
        has_valid_memo = cleaned_memo_check != ""
        
        # 4. 拼接！只对有有效备注的行进行拼接
        # 使用 raw_memo (原始备注) 还是 cleaned_memo_check? 
        # 建议用 raw_memo，因为最后还会统一洗一次，这样能保留用户原始输入
        df.loc[has_valid_memo, '描述'] = df.loc[has_valid_memo, '描述'] + "." + raw_memo[has_valid_memo]

    # 5. 最终统一清洗 (去掉整个描述里的垃圾字符)
    # 这样 "康福路店./" 就会变成 "康福路店"
    df['描述'] = df['描述'].astype(str).str.replace(PATTERNS['CLEAN'], '', regex=True).str.strip()
    # 去掉可能产生的尾部点号 (例如 "康福路店.")
    df['描述'] = df['描述'].str.rstrip('.')

    df['记录类型'] = df.get('记录类型', pd.NA).replace("", pd.NA).fillna(df['收/支'])
    df['商家'] = df.get('商家', pd.NA).replace("", pd.NA).fillna(df['交易对方'])

    df = ensure_columns(df, main_col, sub_col)

    # 餐饮时段识别
    df[sub_col] = df[sub_col].astype(str).str.strip().replace("nan", "")
    uncat_mask = (df[sub_col] == "")
    
    is_meal = (
        df['商家(old)'].str.contains(PATTERNS['MEAL'], regex=True, na=False) | 
        df['商品'].str.contains(PATTERNS['MEAL'], regex=True, na=False) |
        df['商家'].str.contains(PATTERNS['MEAL'], regex=True, na=False) 
    ) & uncat_mask
    
    if is_meal.any():
        h = df.loc[is_meal, '交易时间'].dt.hour
        vals = np.select([(h>=6)&(h<10), (h>=11)&(h<16), (h>=16)&(h<21)], ["早餐", "午餐", "晚餐"], default="夜宵")
        df.loc[is_meal, sub_col] = vals

    df = process_heuristics(df, main_col, sub_col)
    
    raw_amt = df.get('金额(元)', 0)
    df['金额'] = pd.to_numeric(raw_amt, errors='coerce').fillna(0)
    neg_mask = df['记录类型'].isin(['支出','应付款项','应收款项']) & (df.get('收/支')=='支出')
    df.loc[neg_mask, '金额'] *= -1
    
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
    print(f"{bcolors.BOLD}=== Moze 导入工具 v10.25 (店铺+备注拼接版) ==={bcolors.ENDC}")
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

        mask_rec = df_final['记录类型'] == '应收款项'
        invalid_rec = mask_rec & (
            (df_final[main_col] == "") | (df_final[main_col].isna()) |
            (df_final[sub_col] == "") | (df_final[sub_col].isna()) |
            (df_final['对象'] == "") | (df_final['对象'].isna())
        )
        
        if invalid_rec.any():
            print(f"{bcolors.FAIL}[严重] {invalid_rec.sum()} 条【应收款项】记录缺失分类或对象:{bcolors.ENDC}")
            print(df_final.loc[invalid_rec, ['日期', '商家', '描述', '金额', '对象']].head(5).to_string())
        else:
            print(f"{bcolors.OKGREEN}√ 应收款项完美{bcolors.ENDC}")

    except Exception: traceback.print_exc()

if __name__ == "__main__": main(); input("\n回车退出...")