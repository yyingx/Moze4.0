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
#      Moze 导入脚本 v5.8 (纯净水独立版)
# ==========================================

# --- 终端颜色 ---
class bcolors:
    OKGREEN = '\033[92m'; WARNING = '\033[93m'; FAIL = '\033[91m'; ENDC = '\033[0m'; BOLD = '\033[1m'

# --- 路径 ---
RULE_BOOK_PATH = Path(r"E:\天之逸2025\Moze4.0\Moze Dict.xlsx")
TARGET_DIR = Path(r"E:\天之逸2025\Moze4.0\Moze4.0_Import")

# --- 默认配置 ---
CONFIG = {
    'TRANSFER_TARGET_1': '肖恩',
    'TRANSFER_TARGET_2': '工商银行(9579)',
    'CARD_PINGAN': '平安银行信用卡(4946)',
    'CARD_ICBC': '工商银行储蓄卡(9579)',
    'KEYWORD_CHARGING': '自助服务-充电桩',
    'KEYWORD_ALIPAY_CSV': '支付宝支付科技有限公司',
    'KEYWORD_WECHAT_XLSX': '微信支付账单明细',
}

# --- 列映射 ---
COLUMN_MAPPING = {
    '交易时间': '交易时间', '交易类型': '交易类型', '交易对方': '交易对方',
    '商品': '商品', '收/支': '收/支', '金额(元)': '金额(元)', '金额（元）': '金额(元)',
    '支付方式': '支付方式', '当前状态': '当前状态', '备注': '备注',
    '交易单号': '交易单号', '商户单号': '商户单号',
    '交易分类': '交易类型', '商品说明': '商品', '金额': '金额(元)',
    '收/付款方式': '支付方式', '交易状态': '当前状态', '交易订单号': '交易单号',
    '商家订单号': '商户单号',
    '付款时间': '交易时间', '商品名称': '商品', '类型': '交易类型',
}

# ==========================================
#               工具函数
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
    print("正在加载规则...", flush=True)
    try:
        xl = pd.ExcelFile(rule_path, engine='openpyxl')
        sheet_name = None
        for s in xl.sheet_names:
            if s.lower() == 'settings': continue
            tmp = pd.read_excel(rule_path, sheet_name=s, nrows=5, engine='openpyxl')
            if '商家(old)' in tmp.columns: sheet_name = s; break
        
        if not sheet_name: print("错：没找到规则表"); return None
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
        for e in ['utf-8', 'GBK']:
            try:
                with open(file_path, 'r', encoding=e, errors='ignore') as f:
                    for i, l in enumerate(f):
                        if i>50: break
                        if CONFIG['KEYWORD_ALIPAY_CSV'] in l: ftype, skip, enc = "AliPay", i+1, e; break
                if ftype: break
            except: pass
    elif file_path.suffix.lower() == '.xlsx':
        try:
            tmp = pd.read_excel(file_path, header=None, nrows=20, engine='openpyxl')
            if CONFIG['KEYWORD_WECHAT_XLSX'] in str(tmp.iloc[0,0]): ftype, skip = "WeChat", 16
        except: pass
    
    if not ftype: print("  跳过：格式未知"); return None

    try:
        if ftype == "WeChat": df = pd.read_excel(file_path, skiprows=skip, engine='openpyxl')
        else: df = pd.read_csv(file_path, skiprows=skip, encoding=enc)
        df.columns = df.columns.str.strip()
        df.rename(columns=COLUMN_MAPPING, inplace=True)
        if '金额(元)' not in df.columns: df['金额(元)'] = 0
        df['金额(元)'] = pd.to_numeric(df['金额(元)'].astype(str).str.replace('¥','').str.strip(), errors='coerce').fillna(0)
        df['_source_tag'] = '#WechatPay' if ftype == "WeChat" else '#AliPay'
        return df
    except Exception as e: print(f"  加载错: {e}"); return None

# ==========================================
#               业务逻辑
# ==========================================

def ensure_columns(df, main_col, sub_col):
    required_cols = [main_col, sub_col, '金额', '记录类型', '账户', '商家', '描述', '项目', '名称', '对象', '标签', '日期', '时间', '币种', '手续费', '折扣']
    for c in required_cols:
        if c not in df.columns: df[c] = pd.NA
    str_cols = ['记录类型', '项目', '对象', main_col, sub_col, '商家']
    for c in str_cols:
        df[c] = df[c].fillna("")
    return df

def process_transfers(df, main_col, sub_col):
    res = []
    target = CONFIG['TRANSFER_TARGET_1']
    
    mask_in = (df["交易对方"] == target) & (df["收/支"] == "收入")
    if mask_in.any():
        tmp = df[mask_in].copy()
        out_d = tmp.copy(); out_d['金额'] = pd.to_numeric(tmp['金额(元)'])*-1; out_d['记录类型'] = '转出'; out_d['账户'] = '零钱2'
        in_d = tmp.copy(); in_d['金额'] = pd.to_numeric(tmp['金额(元)']); in_d['记录类型'] = '转入'; in_d['账户'] = '零钱3'
        res.extend([out_d, in_d])
        
    mask_out = (df["交易对方"] == target) & (df["收/支"] == "支出")
    if mask_out.any():
        tmp = df[mask_out].copy()
        out_d = tmp.copy(); out_d['金额'] = pd.to_numeric(tmp['金额(元)'])*-1; out_d['记录类型'] = '转出'; out_d['账户'] = '零钱3'
        in_d = tmp.copy(); in_d['金额'] = pd.to_numeric(tmp['金额(元)']); in_d['记录类型'] = '转入'; in_d['账户'] = '零钱2'
        res.extend([out_d, in_d])

    target2 = CONFIG['TRANSFER_TARGET_2']
    mask_icbc = (df["交易对方"] == target2)
    if mask_icbc.any():
        tmp = df[mask_icbc].copy()
        out_d = tmp.copy(); out_d['金额'] = pd.to_numeric(tmp['金额(元)'])*-1; out_d['记录类型'] = '转出'; out_d['账户'] = '零钱3'; out_d[sub_col] = '提现'
        in_d = tmp.copy(); in_d['金额'] = pd.to_numeric(tmp['金额(元)']); in_d['记录类型'] = '转入'; in_d['账户'] = '工商银行'; in_d[sub_col] = '提现'
        res.extend([out_d, in_d])

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
    
    def check_keywords(pattern):
        return (
            df['商家'].str.contains(pattern, regex=True, na=False) |
            df['商品'].str.contains(pattern, regex=True, na=False) |
            df['描述'].str.contains(pattern, regex=True, na=False)
        )

    uncat = (df[sub_col] == "")
    
    # 允许覆写的“弱分类”：空值 或 食材
    allow_overwrite = (df[sub_col] == "") | (df[sub_col] == "食材")

    # 3. 充电桩
    mask = df['商品'].str.contains(CONFIG['KEYWORD_CHARGING'], na=False) & uncat
    if mask.any(): 
        df.loc[mask, ['名称', main_col, sub_col]] = ['充电', '交通', '加油充电']
        split = df.loc[mask, '商品'].str.split('/', n=1, expand=True)
        if not split.empty and split.shape[1]>1: df.loc[mask, '描述'] = split[1].str.strip()
    
    # 4. 软件
    mask = check_keywords("软件") & uncat
    if mask.any(): df.loc[mask, [sub_col, main_col, '项目']] = ['Software', '虚拟', '学习']
    
    # 5. 水果
    pat = r"(?:水果|苹果|香蕉|红枣|大枣|桃子|水蜜桃|樱桃|黄桃|香梨|雪梨|柑橘|沃柑|草莓|甘蔗|桔子|沙糖桔|百果园|鲜果|果业|葡萄|哈密瓜|西瓜|甜瓜|柚子|柠檬|橙子|菠萝|凤梨|榴莲|山竹|蓝莓|荔枝|龙眼|椰子|柿子|杨梅|李子|芒果|猕猴桃|火龙果|百香果|莲雾|车厘子)"
    mask = check_keywords(pat) & (~mask) & allow_overwrite
    if mask.any(): df.loc[mask, ['名称', sub_col, main_col, '项目']] = ['水果', '饮料水果', '饮食', '食']
    
    # 6.1 [独立] 纯净水/矿泉水 -> 纯净水
    pat_water = r"(?:纯净水|矿泉水)"
    mask = check_keywords(pat_water) & (~mask) & allow_overwrite
    if mask.any():
        # 子类别设为您专用的 '纯净水'
        df.loc[mask, ['名称', sub_col, main_col, '项目']] = ['纯净水', '纯净水', '饮食', '食']

    # 6.2 其他饮料 (剔除了水)
    pat = r"(?:可乐|红牛|奶茶|东鹏|饮料|汽水|果汁|椰汁|酸奶)"
    mask = check_keywords(pat) & (~mask) & allow_overwrite
    if mask.any(): df.loc[mask, ['名称', sub_col, main_col, '项目']] = ['饮料', '饮料水果', '饮食', '食']

    # 7. 馒头 -> 食材
    mask = check_keywords("馒头") & uncat
    if mask.any():
        df.loc[mask, ['名称', sub_col, main_col]] = ['馒头', '食材', '饮食']
        df[sub_col] = df[sub_col].replace("nan", "")

    # 8. 餐饮/食品类项目补全
    food_list = ["午餐", "纯净水", "饮料水果", "早餐", "晚餐", "零食", "液化气费", "夜宵", "食材"]
    mask_food = df[sub_col].isin(food_list) & (df['项目'] == "")
    if mask_food.any():
        df.loc[mask_food, '项目'] = "食"
    
    # 9. 垃圾信息清理
    pat = r"(?:支付|商户|二维码收款|收款|美团|编号|转账|错误|订单编号|\d{10,})"
    mask = df['描述'].str.contains(pat, regex=True)
    if '备注' in df.columns: mask |= df['备注'].astype(str).str.contains(r'(?:\d{10,}|订单编号)', na=False)
    df.loc[mask, '描述'] = ""
    
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
            
    if '描述' in df.columns and '备注' in df.columns:
        mask = (df['_source_tag']=='#AliPay') & df['描述'].notna() & df['备注'].notna()
        if mask.any(): df.loc[mask, '描述'] = df.loc[mask, '描述'] + '.' + df.loc[mask, '备注']

    df['描述'] = df.get('描述', pd.NA).fillna(df.get('商品', pd.NA))
    df['记录类型'] = df.get('记录类型', pd.NA).replace("", pd.NA).fillna(df['收/支'])
    df['商家'] = df.get('商家', pd.NA).replace("", pd.NA).fillna(df['交易对方'])

    df = ensure_columns(df, main_col, sub_col)

    # 餐饮时段识别
    df[sub_col] = df[sub_col].astype(str).str.strip().replace("nan", "")
    uncat_mask = (df[sub_col] == "")
    
    pat = r"(?:三镇民生|兰州|丝路|沙县|永和四喜|水煎包|老乡鸡|混沌|馄饨|牛肉汤|热干面|黄蜀郎|鸡公煲|小吃|烧烤|食堂|路边摊|麦香园|长沙臭豆腐|东苑一层|东苑二层|西区食堂|包子|小笼包|外勤|板面|水饺|猪脚饭|肠粉|卤肉饭|油泼面|重庆小面|牛杂粉|麻辣香锅|麻辣烫|煎包|煎饼)"
    
    is_meal = (
        df['商家(old)'].str.contains(pat, regex=True, na=False) | 
        df['商品'].str.contains(pat, regex=True, na=False) |
        df['商家'].str.contains(pat, regex=True, na=False) 
    ) & uncat_mask
    
    if is_meal.any():
        h = df.loc[is_meal, '交易时间'].dt.hour
        vals = np.select([(h>=6)&(h<10), (h>=11)&(h<16), (h>=16)&(h<21)], ["早餐", "午餐", "晚餐"], default="夜宵")
        df.loc[is_meal, main_col] = '饮食'; df.loc[is_meal, sub_col] = vals; df.loc[is_meal, '项目'] = '食'

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

# ==========================================
#               主程序
# ==========================================

def main():
    print(f"{bcolors.BOLD}=== Moze 导入工具 v5.8 (纯净水独立) ==={bcolors.ENDC}")
    try:
        load_settings(RULE_BOOK_PATH)
        df_rules = load_rules(RULE_BOOK_PATH)
        if df_rules is None: input("按 Enter 退出"); return

        root = tk.Tk(); root.withdraw()
        print("\n选择文件...")
        files = filedialog.askopenfilenames(filetypes=(("Excel/CSV", "*.xlsx;*.csv"), ("All", "*.*")))
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
        
        # 1. 支出
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

        # 2. 应收款项
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