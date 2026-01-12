from nba_api.stats.endpoints import scoreboardv2, boxscoretraditionalv3
from rich.console import Console
from rich.table import Table
from rich.text import Text
from rich import box
import pandas as pd
import time

# =================配置区域=================
TARGET_DATE = "2025-11-20"  # 目标日期

# 样式配置
STYLE_HEADER_BG = "bold white on blue"
STYLE_TITLE = "bold green"
STYLE_TEXT = "cyan"
STYLE_DNP = "dim cyan"
STYLE_SCOREBOARD = "bold white on blue"
# =========================================

console = Console()

def calculate_eff(row):
    """计算效率值 (适配 V3 数据结构)"""
    try:
        # V3 的分钟数通常是 "PT32M10S" 或 "32:10"，如果没有时间则为 None
        if pd.isna(row.get('MIN')) or row.get('MIN') == "": return 0
        
        missed_fg = row['FGA'] - row['FGM']
        missed_ft = row['FTA'] - row['FTM']
        eff = (row['PTS'] + row['REB'] + row['AST'] + 
               row['STL'] + row['BLK']) - (missed_fg + missed_ft + row['TO'])
        return int(eff)
    except:
        return 0

def fmt_ma(made, attempted):
    return f"{int(made)}-{int(attempted)}"

def fmt_pm(val):
    if pd.isna(val): return "0"
    val = int(val)
    if val > 0: return f"[green]+{val}[/green]"
    if val < 0: return f"[red]{val}[/red]"
    return str(val)

def normalize_v3_player_data(players_dict_list):
    """核心函数：将 V3 的嵌套 JSON 转换为扁平的 DataFrame"""
    if not players_dict_list:
        return pd.DataFrame()
        
    # 使用 pandas 自动展平嵌套的 statistics 字段
    df = pd.json_normalize(players_dict_list)
    
    # V3 的列名通常是 statistics.points 这样的格式
    # 我们需要把它们重命名为我们习惯的简写，方便后续处理
    rename_map = {
        'firstName': 'FIRST_NAME',
        'familyName': 'LAST_NAME',
        'position': 'POSITION',
        'statistics.minutes': 'MIN',
        'statistics.fieldGoalsMade': 'FGM',
        'statistics.fieldGoalsAttempted': 'FGA',
        'statistics.threePointersMade': 'FG3M',
        'statistics.threePointersAttempted': 'FG3A',
        'statistics.freeThrowsMade': 'FTM',
        'statistics.freeThrowsAttempted': 'FTA',
        'statistics.reboundsOffensive': 'OREB',
        'statistics.reboundsTotal': 'REB',
        'statistics.assists': 'AST',
        'statistics.steals': 'STL',
        'statistics.blocks': 'BLK',
        'statistics.turnovers': 'TO',
        'statistics.foulsPersonal': 'PF',
        'statistics.points': 'PTS',
        'statistics.plusMinusPoints': 'PLUS_MINUS'
    }
    
    # 只保留存在的列进行重命名
    cols_to_rename = {k: v for k, v in rename_map.items() if k in df.columns}
    df = df.rename(columns=cols_to_rename)
    
    # 合并名字
    df['PLAYER_NAME'] = df['FIRST_NAME'] + " " + df['LAST_NAME']
    
    # 填充空值 (V3 有时会将 0 返回为 NaN)
    numeric_cols = ['FGM', 'FGA', 'FG3M', 'FG3A', 'FTM', 'FTA', 'OREB', 'REB', 'AST', 'STL', 'BLK', 'TO', 'PF', 'PTS', 'PLUS_MINUS']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].fillna(0)
            
    return df

def create_retro_table(players_df):
    """创建复古风格表格"""
    table = Table(show_header=True, header_style=STYLE_HEADER_BG, box=None, padding=(0, 1), collapse_padding=True, style=STYLE_TEXT)
    
    cols = ["MIN", "FGM-A", "3PM-A", "FTM-A", "+/-", "OFF", "REB", "AST", "STL", "BS", "TO", "PF", "PTS", "EFF"]
    table.add_column("STARTERS", justify="left", style=STYLE_TEXT, no_wrap=True, width=20)
    for col in cols:
        table.add_column(col, justify="right", style=STYLE_TEXT)

    # V3 处理分钟数格式: 可能是 "PT12M00.00S" 也可能是 "12:00"
    def format_min(m):
        if pd.isna(m) or m == "": return None
        m = str(m)
        if "PT" in m: # ISO 格式处理
            import re
            match = re.search(r'PT(\d+)M', m)
            return match.group(1) if match else "0"
        return m.split(':')[0] # 普通格式 "28:15" -> "28"

    # 分离首发替补
    # 在 V3 中，未上场球员的 MIN 通常为空
    active = players_df[players_df['MIN'].notna() & (players_df['MIN'] != "")]
    dnp = players_df[players_df['MIN'].isna() | (players_df['MIN'] == "")]
    
    # V3 也可以简单地通过是否有 MIN 时间来粗略判断，或者看 position
    # 这里假设有位置且有时间的是首发（不太严谨但视觉上够用），或者简单的前5个
    # 更严谨的是读取 starter 字段，但 json_normalize 后可能要找一下
    # 为了通用，我们这里简单逻辑：前5个有时间的是首发 (API 返回顺序通常是首发在前)
    
    starters = active.head(5)
    bench = active.iloc[5:]

    def add_rows(df):
        for _, p in df.iterrows():
            name_display = p['PLAYER_NAME']
            if p.get('POSITION'): name_display += f", {p['POSITION']}"
            
            table.add_row(
                name_display,
                format_min(p['MIN']),
                fmt_ma(p['FGM'], p['FGA']),
                fmt_ma(p['FG3M'], p['FG3A']),
                fmt_ma(p['FTM'], p['FTA']),
                fmt_pm(p['PLUS_MINUS']),
                str(int(p['OREB'])),
                str(int(p['REB'])),
                str(int(p['AST'])),
                str(int(p['STL'])),
                str(int(p['BLK'])),
                str(int(p['TO'])),
                str(int(p['PF'])),
                f"[bold yellow]{int(p['PTS'])}[/]" if p['PTS'] >= 20 else str(int(p['PTS'])),
                str(calculate_eff(p))
            )

    add_rows(starters)
    table.add_row(Text("BENCH", style=STYLE_HEADER_BG), *["" for _ in cols], style=STYLE_HEADER_BG)
    add_rows(bench)

    for _, p in dnp.iterrows():
        table.add_row(p['PLAYER_NAME'], Text("DNP - Coach's Decision", style=STYLE_DNP, justify="center"), *[""]*(len(cols)-1))

    return table

def process_game(game_id, summary):
    matchup = summary['GAMECODE']
    print(f"正在下载数据: {matchup} ...")
    
    try:
        # === 使用 V3 接口 ===
        box = boxscoretraditionalv3.BoxScoreTraditionalV3(game_id=game_id)
        # V3 必须用 get_dict() 获取原始字典
        data = box.get_dict()
        
        box_data = data.get('boxScoreTraditional', {})
        home_team_raw = box_data.get('homeTeam', {})
        away_team_raw = box_data.get('awayTeam', {})
        
        # 解析两队球员数据
        home_players_df = normalize_v3_player_data(home_team_raw.get('players', []))
        away_players_df = normalize_v3_player_data(away_team_raw.get('players', []))
        
        # 打印分割线
        console.print("="*60, style="blue")
        
        # 1. 比分板 (V3 字典里也有比分)
        score_table = Table(show_header=True, header_style=STYLE_SCOREBOARD, box=None)
        score_table.add_column("TEAMS", style="bold white")
        score_table.add_column("SCORE", justify="right", style="bold white")
        
        # 从 raw data 获取队名和比分
        score_table.add_row(f"{away_team_raw.get('teamCity')} {away_team_raw.get('teamName')}", str(away_team_raw.get('score')))
        score_table.add_row(f"{home_team_raw.get('teamCity')} {home_team_raw.get('teamName')}", str(home_team_raw.get('score')))
        console.print(score_table)
        console.print()

        # 2. 客队数据
        console.print(Text(f"{away_team_raw.get('teamCity')} {away_team_raw.get('teamName')}", style=STYLE_TITLE))
        console.print(create_retro_table(away_players_df))
        console.print()

        # 3. 主队数据
        console.print(Text(f"{home_team_raw.get('teamCity')} {home_team_raw.get('teamName')}", style=STYLE_TITLE))
        console.print(create_retro_table(home_players_df))
        
        print("\n\n")
        
    except Exception as e:
        print(f"❌ 处理比赛 {game_id} 时出错: {e}")
        # 打印详细错误堆栈以便调试
        # import traceback
        # traceback.print_exc()

def main():
    print(f"🏀 正在获取 {TARGET_DATE} 的所有比赛列表...")
    board = scoreboardv2.ScoreboardV2(game_date=TARGET_DATE)
    games = board.game_header.get_data_frame()
    
    if games.empty:
        print("❌ 该日期没有比赛。")
        return

    print(f"✅ 找到 {len(games)} 场比赛。开始处理...\n")
    
    for _, game in games.iterrows():
        process_game(game['GAME_ID'], game)
        time.sleep(1.0)

if __name__ == "__main__":
    main()