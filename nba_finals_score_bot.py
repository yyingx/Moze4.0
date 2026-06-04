import argparse
import datetime as dt
import getpass
from html.parser import HTMLParser
import json
import os
import re
import sys
from pathlib import Path
from typing import Any

import requests


ESPN_SCOREBOARD = "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/scoreboard"
ESPN_SUMMARY_ENDPOINTS = [
    "https://site.web.api.espn.com/apis/site/v2/sports/basketball/nba/summary",
    "https://site.api.espn.com/apis/site/v2/sports/basketball/nba/summary",
]
YAHOO_SCOREBOARD = "https://sports.yahoo.com/nba/scoreboard/"
NEWSMTH_BASE = "https://www.newsmth.net/nForum"
NEWSMTH_MOBILE_LOGIN = "https://m.newsmth.net/user/login"


class SportsHTMLParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.text_parts: list[str] = []
        self.links: list[str] = []
        self.tables: list[list[list[str]]] = []
        self._table: list[list[str]] | None = None
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attrs_dict = {key: value or "" for key, value in attrs}
        if tag == "a" and attrs_dict.get("href"):
            self.links.append(attrs_dict["href"])
        if tag in ("p", "div", "section", "article", "br", "tr", "li", "h1", "h2", "h3"):
            self.text_parts.append("\n")
        if tag == "table":
            self._table = []
        elif tag == "tr" and self._table is not None:
            self._row = []
        elif tag in ("th", "td") and self._row is not None:
            self._cell = []

    def handle_data(self, data: str) -> None:
        if data:
            self.text_parts.append(data)
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag in ("th", "td") and self._cell is not None and self._row is not None:
            cell = " ".join(" ".join(self._cell).split())
            self._row.append(cell)
            self._cell = None
        elif tag == "tr" and self._row is not None and self._table is not None:
            if any(cell for cell in self._row):
                self._table.append(self._row)
            self._row = None
        elif tag == "table" and self._table is not None:
            if self._table:
                self.tables.append(self._table)
            self._table = None

    @property
    def text(self) -> str:
        raw = "".join(self.text_parts)
        return "\n".join(line.strip() for line in raw.splitlines() if line.strip())


def parse_html(html: str) -> SportsHTMLParser:
    parser = SportsHTMLParser()
    parser.feed(html)
    return parser


def date_window(center: dt.date, days_back: int, days_forward: int) -> list[str]:
    dates = []
    for offset in range(-days_back, days_forward + 1):
        day = center + dt.timedelta(days=offset)
        dates.append(day.strftime("%Y%m%d"))
    return dates


def fetch_scoreboard(date_text: str) -> dict[str, Any]:
    response = requests.get(
        ESPN_SCOREBOARD,
        params={"dates": date_text},
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=15,
    )
    response.raise_for_status()
    return response.json()


def fetch_summary(event_id: str) -> dict[str, Any]:
    last_error: Exception | None = None
    for endpoint in ESPN_SUMMARY_ENDPOINTS:
        try:
            response = requests.get(
                endpoint,
                params={"event": event_id},
                headers={"User-Agent": "Mozilla/5.0"},
                timeout=15,
            )
            response.raise_for_status()
            return response.json()
        except requests.RequestException as exc:
            last_error = exc
    raise RuntimeError(f"Could not fetch ESPN boxscore summary for event {event_id}: {last_error}")


def yahoo_date(date_text: str) -> str:
    if "-" in date_text:
        return date_text
    value = dt.datetime.strptime(date_text, "%Y%m%d").date()
    return value.strftime("%Y-%m-%d")


def fetch_yahoo_html(url: str, params: dict[str, str] | None = None) -> str:
    response = requests.get(
        url,
        params=params,
        headers={
            "User-Agent": "Mozilla/5.0",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        },
        timeout=20,
    )
    response.raise_for_status()
    return response.text


def normalize_yahoo_url(href: str) -> str:
    if href.startswith("//"):
        return "https:" + href
    if href.startswith("/"):
        return "https://sports.yahoo.com" + href
    return href


def yahoo_game_urls(date_text: str) -> list[str]:
    html = fetch_yahoo_html(YAHOO_SCOREBOARD, {"date": yahoo_date(date_text)})
    parser = parse_html(html)
    urls: list[str] = []
    seen: set[str] = set()
    for href in parser.links:
        url = normalize_yahoo_url(href).split("?")[0]
        if not re.search(r"^https://sports\.yahoo\.com/nba/[a-z0-9-]+-\d{10}/?$", url):
            continue
        if url not in seen:
            seen.add(url)
            urls.append(url)
    return urls


def yahoo_score_rows(tables: list[list[list[str]]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for table in tables:
        if not table:
            continue
        header = [cell.lower() for cell in table[0]]
        if "tot" not in header and "total" not in header:
            continue
        for row in table[1:]:
            if len(row) < 2:
                continue
            score = row[-1]
            if not re.fullmatch(r"\d+|-", score):
                continue
            team = row[0].strip()
            if team:
                rows.append({"team": team, "score": score, "columns": table[0], "values": row})
        if len(rows) >= 2:
            return rows[:2]
    return rows


def yahoo_status(text: str) -> str:
    for label in ("Final/OT", "Final", "Halftime", "Scheduled", "Postponed"):
        if label.lower() in text.lower():
            return label
    match = re.search(r"\b(?:[1-4](?:st|nd|rd|th)|OT)\s+\d{1,2}:\d{2}\b", text)
    return match.group(0) if match else "NBA"


def yahoo_player_sections(
    tables: list[list[list[str]]],
    score_rows: list[dict[str, Any]],
    max_players_per_team: int,
) -> list[str]:
    stat_tables: list[list[list[str]]] = []
    for table in tables:
        if not table:
            continue
        header = [cell.upper() for cell in table[0]]
        if "MIN" in header and "PTS" in header and "FG" in header:
            stat_tables.append(table)

    sections: list[str] = []
    stat_fields = ["MIN", "PTS", "REB", "AST", "STL", "BLK", "TO", "FG", "3PT", "FT", "+/-"]

    for index, table in enumerate(stat_tables[:2]):
        labels = [cell.upper() for cell in table[0]]
        rows: list[tuple[int, str]] = []
        for row in table[1:]:
            if len(row) < len(labels):
                continue
            name = row[0].strip()
            if not name or name.lower() in ("starters", "bench", "team totals"):
                continue
            mapped = {label: row[pos] for pos, label in enumerate(labels) if pos < len(row)}
            if not has_played(mapped):
                continue
            try:
                points = int(mapped.get("PTS", "0"))
            except ValueError:
                points = 0
            rows.append((points, player_line(name, mapped, stat_fields)))

        if not rows:
            continue
        rows.sort(key=lambda item: item[0], reverse=True)
        team_name_text = score_rows[index]["team"] if index < len(score_rows) else f"Team {index + 1}"
        sections.append("\n".join([f"{team_name_text} player stats:", *[line for _, line in rows[:max_players_per_team]]]))

    return sections


def stat_tables(tables: list[list[list[str]]]) -> list[list[list[str]]]:
    result: list[list[list[str]]] = []
    for table in tables:
        if not table:
            continue
        header = [cell.upper() for cell in table[0]]
        if "MIN" in header and "PTS" in header and "FG" in header:
            result.append(table)
    return result


def table_line(widths: list[int], left: str, mid: str, right: str, fill: str = "─") -> str:
    return left + mid.join(fill * width for width in widths) + right


def table_row(values: list[str], widths: list[int]) -> str:
    cells = []
    for value, width in zip(values, widths):
        text = str(value)
        cells.append(text[:width].ljust(width))
    return "│" + "│".join(cells) + "│"


def stat_value(mapped: dict[str, str], *names: str) -> str:
    for name in names:
        value = mapped.get(name.upper())
        if value not in (None, ""):
            return value
    return "-"


def format_scoreboard(score_rows: list[dict[str, Any]]) -> str:
    if len(score_rows) < 2:
        return "Scoreboard unavailable."

    columns = [str(cell) for cell in score_rows[0].get("columns", [])]
    keep_indexes = list(range(1, len(columns)))
    labels = [columns[i] for i in keep_indexes]
    if labels:
        labels[-1] = "T"
    rows = []
    for row in score_rows[:2]:
        values = [str(cell) for cell in row.get("values", [])]
        team = str(row.get("team", ""))[:18]
        period_values = [values[i] if i < len(values) else "-" for i in keep_indexes]
        rows.append([team, *period_values])

    widths = [18, *[max(3, len(label)) for label in labels]]
    output = [
        table_line(widths, "┌", "┬", "┐"),
        table_row(["TEAM", *labels], widths),
        table_line(widths, "├", "┼", "┤"),
        *[table_row(row, widths) for row in rows],
        table_line(widths, "└", "┴", "┘"),
    ]
    return "\n".join(output)


def format_ptt_player_table(team_name_text: str, table: list[list[str]], max_players: int) -> str:
    labels = [cell.upper() for cell in table[0]]
    rows: list[tuple[int, list[str]]] = []
    for row in table[1:]:
        if len(row) < len(labels):
            continue
        name = row[0].strip()
        if not name or name.lower() in ("starters", "bench", "team totals"):
            continue
        mapped = {label: row[pos] for pos, label in enumerate(labels) if pos < len(row)}
        if not has_played(mapped):
            continue
        try:
            points = int(stat_value(mapped, "PTS").replace("-", "0"))
        except ValueError:
            points = 0
        display = [
            name,
            stat_value(mapped, "MIN"),
            stat_value(mapped, "FG"),
            stat_value(mapped, "3PT", "3P"),
            stat_value(mapped, "FT"),
            stat_value(mapped, "+/-", "+-"),
            stat_value(mapped, "OREB", "OR"),
            stat_value(mapped, "DREB", "DR"),
            stat_value(mapped, "REB", "TR"),
            stat_value(mapped, "AST", "A"),
            stat_value(mapped, "PF"),
            stat_value(mapped, "STL", "ST"),
            stat_value(mapped, "TO"),
            stat_value(mapped, "BLK", "BS"),
            stat_value(mapped, "PTS"),
        ]
        rows.append((points, display))

    rows.sort(key=lambda item: item[0], reverse=True)
    rows = rows[:max_players]
    header = ["PLAYER", "MIN", "FG", "3PT", "FT", "+/-", "OR", "DR", "TR", "A", "PF", "ST", "TO", "BS", "PTS"]
    widths = [18, 5, 7, 7, 7, 5, 3, 3, 3, 3, 3, 3, 3, 3, 4]
    output = [
        team_name_text,
        table_line(widths, "┌", "┬", "┐"),
        table_row(header, widths),
        table_line(widths, "├", "┼", "┤"),
        *[table_row(row, widths) for _, row in rows],
        table_line(widths, "└", "┴", "┘"),
    ]
    return "\n".join(output)


def build_yahoo_ptt_post(
    game: dict[str, Any],
    max_players_per_team: int,
    include_player_stats: bool = True,
) -> tuple[str, str, str]:
    text = str(game.get("text") or "")
    tables = game.get("tables") or []
    score_rows = yahoo_score_rows(tables)
    status = yahoo_status(text)

    if len(score_rows) >= 2:
        away, home = score_rows[0], score_rows[1]
        subject = f"[BOX ] {away['team']} {away['score']}:{home['score']} {home['team']}"
    else:
        subject = f"[BOX ] NBA Finals {status}"

    parts = [
        "NBA Finals Box Score",
        f"Status: {status}",
        f"Source: Yahoo Sports ({game.get('url')})",
        "",
        format_scoreboard(score_rows),
    ]

    if include_player_stats:
        stats = stat_tables(tables)
        if stats:
            for index, table in enumerate(stats[:2]):
                team_name_text = score_rows[index]["team"] if index < len(score_rows) else f"Team {index + 1}"
                parts.extend(["", format_ptt_player_table(team_name_text, table, max_players_per_team)])
        else:
            parts.extend(["", "Player stats: not available yet from Yahoo page."])

    state_key = "|".join(
        [
            str(game.get("url") or ""),
            status,
            "|".join(f"{row['team']}:{row['score']}" for row in score_rows),
            "ptt",
        ]
    )
    return subject[:80], "\n".join(parts), state_key


def pick_yahoo_event(dates: list[str], allow_any: bool) -> dict[str, Any]:
    all_games: list[dict[str, Any]] = []
    finals_games: list[dict[str, Any]] = []
    for date_text in dates:
        for url in yahoo_game_urls(date_text):
            html = fetch_yahoo_html(url)
            parser = parse_html(html)
            game = {
                "source_date": date_text,
                "url": url,
                "text": parser.text,
                "tables": parser.tables,
            }
            all_games.append(game)
            if "final" in parser.text.lower():
                finals_games.append(game)

    if finals_games:
        return finals_games[0]
    if allow_any and all_games:
        return all_games[0]
    searched = ", ".join(dates)
    raise RuntimeError(f"No Yahoo NBA Finals game found for dates: {searched}")


def build_yahoo_post(
    game: dict[str, Any],
    max_players_per_team: int,
    include_player_stats: bool = True,
) -> tuple[str, str, str]:
    text = str(game.get("text") or "")
    tables = game.get("tables") or []
    score_rows = yahoo_score_rows(tables)
    status = yahoo_status(text)

    if len(score_rows) >= 2:
        away, home = score_rows[0], score_rows[1]
        subject = f"[NBA Finals] {status}: {away['team']} {away['score']}-{home['score']} {home['team']}"
        score_block = "\n".join([f"{away['team']} {away['score']}", f"{home['team']} {home['score']}"])
    else:
        subject = f"[NBA Finals] {status}"
        score_block = "Team score: unavailable from Yahoo page."

    body_parts = [
        "NBA Finals",
        "",
        status,
        score_block,
        "",
        f"Source: Yahoo Sports ({game.get('url')})",
    ]

    if include_player_stats:
        sections = yahoo_player_sections(tables, score_rows, max_players_per_team)
        if sections:
            body_parts.extend(["", *sections])
        else:
            body_parts.extend(["", "Player stats: not available yet from Yahoo page."])

    state_key = "|".join(
        [
            str(game.get("url") or ""),
            status,
            "|".join(f"{row['team']}:{row['score']}" for row in score_rows),
        ]
    )
    return subject[:80], "\n\n".join(body_parts), state_key


def note_text(competition: dict[str, Any]) -> str:
    notes = competition.get("notes") or []
    parts = []
    for note in notes:
        if isinstance(note, dict):
            parts.append(str(note.get("headline") or note.get("text") or ""))
        else:
            parts.append(str(note))
    return " ".join(part for part in parts if part)


def is_finals_event(event: dict[str, Any]) -> bool:
    competition = (event.get("competitions") or [{}])[0]
    haystack = " ".join(
        [
            str(event.get("name") or ""),
            str(event.get("shortName") or ""),
            str(event.get("season", {}).get("slug") or ""),
            note_text(competition),
        ]
    ).lower()
    return "final" in haystack


def team_label(team: dict[str, Any]) -> str:
    info = team.get("team") or {}
    return str(info.get("abbreviation") or info.get("shortDisplayName") or info.get("displayName") or "?")


def team_name(team: dict[str, Any]) -> str:
    info = team.get("team") or {}
    return str(info.get("displayName") or info.get("shortDisplayName") or team_label(team))


def competitor(event: dict[str, Any], home_away: str) -> dict[str, Any]:
    competition = (event.get("competitions") or [{}])[0]
    for item in competition.get("competitors") or []:
        if item.get("homeAway") == home_away:
            return item
    return {}


def score_text(team: dict[str, Any]) -> str:
    score = team.get("score")
    return str(score) if score not in (None, "") else "-"


def event_status(event: dict[str, Any]) -> dict[str, Any]:
    competition = (event.get("competitions") or [{}])[0]
    return competition.get("status") or event.get("status") or {}


def status_label(status: dict[str, Any]) -> str:
    status_type = status.get("type") or {}
    if status_type.get("completed"):
        return "Final"
    if status_type.get("state") == "pre":
        return "Scheduled"
    detail = status_type.get("shortDetail") or status_type.get("detail") or status_type.get("description")
    return str(detail or "Live")


def event_time(event: dict[str, Any]) -> str:
    raw = event.get("date")
    if not raw:
        return "unknown time"
    try:
        value = dt.datetime.fromisoformat(raw.replace("Z", "+00:00"))
        return value.astimezone().strftime("%Y-%m-%d %H:%M %Z")
    except ValueError:
        return str(raw)


def pick_event(dates: list[str], allow_any: bool) -> tuple[str, dict[str, Any]]:
    candidates: list[tuple[str, dict[str, Any]]] = []
    all_events: list[tuple[str, dict[str, Any]]] = []

    for date_text in dates:
        data = fetch_scoreboard(date_text)
        for event in data.get("events") or []:
            all_events.append((date_text, event))
            if is_finals_event(event):
                candidates.append((date_text, event))

    if candidates:
        return candidates[0]
    if allow_any and all_events:
        return all_events[0]

    searched = ", ".join(dates)
    raise RuntimeError(f"No NBA Finals event found in ESPN scoreboard dates: {searched}")


def build_post(event: dict[str, Any]) -> tuple[str, str, str]:
    competition = (event.get("competitions") or [{}])[0]
    status = event_status(event)
    away = competitor(event, "away")
    home = competitor(event, "home")
    away_score = score_text(away)
    home_score = score_text(home)
    label = status_label(status)
    note = note_text(competition) or "NBA Finals"

    status_type = status.get("type") or {}
    if status_type.get("state") == "pre":
        subject = f"[NBA Finals] {team_label(away)} @ {team_label(home)} - {label}"
    else:
        subject = f"[NBA Finals] {label}: {team_label(away)} {away_score}-{home_score} {team_label(home)}"

    body = "\n".join(
        [
            note,
            "",
            f"{label}",
            f"{team_name(away)} {away_score}",
            f"{team_name(home)} {home_score}",
            "",
            f"Game time: {event_time(event)}",
            "Source: ESPN NBA scoreboard",
        ]
    )

    state_key = "|".join(
        [
            str(event.get("id") or ""),
            label,
            str(status.get("period") or ""),
            str(status.get("displayClock") or ""),
            away_score,
            home_score,
        ]
    )
    return subject[:80], body, state_key


def stat_map(labels: list[str], stats: list[Any]) -> dict[str, str]:
    return {str(label).upper(): str(value) for label, value in zip(labels, stats)}


def has_played(mapped: dict[str, str]) -> bool:
    minutes = mapped.get("MIN", "")
    points = mapped.get("PTS", "")
    return minutes not in ("", "0", "0:00", "DNP") or points not in ("", "0")


def player_line(name: str, mapped: dict[str, str], fields: list[str]) -> str:
    values = " ".join(f"{field}:{mapped.get(field, '-')}" for field in fields)
    return f"{name} {values}"


def boxscore_sections(summary: dict[str, Any], max_players_per_team: int) -> list[str]:
    boxscore = summary.get("boxscore") or {}
    players = boxscore.get("players") or []
    sections: list[str] = []
    fields = ["MIN", "PTS", "REB", "AST", "STL", "BLK", "TO", "FG", "3PT", "FT", "+/-"]

    for team_entry in players:
        team = team_entry.get("team") or {}
        team_name_text = str(team.get("displayName") or team.get("shortDisplayName") or team.get("abbreviation") or "Team")
        stats_groups = team_entry.get("statistics") or []
        if not stats_groups:
            continue

        # Basketball boxscores normally put the game totals in the first group.
        group = stats_groups[0]
        labels = [str(label).upper() for label in (group.get("labels") or group.get("keys") or [])]
        athletes = group.get("athletes") or []
        rows: list[tuple[int, str]] = []

        for item in athletes:
            athlete = item.get("athlete") or {}
            name = str(athlete.get("shortName") or athlete.get("displayName") or athlete.get("name") or "").strip()
            if not name:
                continue
            mapped = stat_map(labels, item.get("stats") or [])
            if not has_played(mapped):
                continue
            try:
                points = int(mapped.get("PTS", "0"))
            except ValueError:
                points = 0
            rows.append((points, player_line(name, mapped, fields)))

        if not rows:
            continue

        rows.sort(key=lambda row: row[0], reverse=True)
        shown = [line for _, line in rows[:max_players_per_team]]
        sections.append("\n".join([f"{team_name_text} player stats:", *shown]))

    return sections


def load_state(path: Path) -> set[str]:
    if not path.exists():
        return set()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return set()
    return set(data.get("posted_keys") or [])


def save_state(path: Path, keys: set[str]) -> None:
    path.write_text(json.dumps({"posted_keys": sorted(keys)}, ensure_ascii=False, indent=2), encoding="utf-8")


def login_newsmth(session: requests.Session, username: str, password: str) -> dict[str, Any]:
    response = session.post(
        f"{NEWSMTH_BASE}/user/ajax_login.json",
        data={"id": username, "passwd": password},
        headers={
            "User-Agent": "Mozilla/5.0",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": "https://www.newsmth.net/",
        },
        timeout=15,
    )
    response.raise_for_status()
    result = response.json()
    if str(result.get("ajax_code")) == "0005":
        return result

    # The desktop endpoint may require a captcha and return 1103. The mobile
    # login path often accepts the same form data and sets shared cookies.
    mobile_response = session.post(
        NEWSMTH_MOBILE_LOGIN,
        data={"id": username, "passwd": password},
        headers={
            "User-Agent": "Mozilla/5.0",
            "Content-Type": "application/x-www-form-urlencoded",
            "Referer": NEWSMTH_MOBILE_LOGIN,
        },
        timeout=15,
        allow_redirects=True,
    )
    mobile_response.raise_for_status()

    session_response = session.get(
        f"{NEWSMTH_BASE}/user/ajax_session.json",
        headers={
            "User-Agent": "Mozilla/5.0",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": "https://www.newsmth.net/",
        },
        timeout=15,
    )
    session_response.raise_for_status()
    try:
        session_result = session_response.json()
    except json.JSONDecodeError:
        session_result = {}

    if session_result.get("is_login") or session.cookies.get("main[UTMPUSERID]"):
        return {
            "ajax_st": 1,
            "ajax_code": "0005",
            "ajax_msg": "操作成功",
            "fallback": "mobile",
            "session": session_result,
        }

    result["mobile_status"] = mobile_response.status_code
    result["mobile_url"] = mobile_response.url
    result["mobile_session"] = session_result
    return result


def post_newsmth(session: requests.Session, board: str, subject: str, content: str) -> dict[str, Any]:
    response = session.post(
        f"{NEWSMTH_BASE}/article/{board}/ajax_post.json",
        data={"id": "0", "subject": subject, "content": content, "signature": "0"},
        headers={
            "User-Agent": "Mozilla/5.0",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": f"https://www.newsmth.net/nForum/#!board/{board}",
        },
        timeout=15,
    )
    response.raise_for_status()
    try:
        return response.json()
    except json.JSONDecodeError:
        return {"raw": response.text[:1000]}


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate or post an NBA Finals score update to NewSMTH.")
    parser.add_argument("--source", choices=["yahoo", "espn"], default="yahoo")
    parser.add_argument("--format", choices=["ptt", "simple"], default="ptt")
    parser.add_argument("--date", help="ESPN scoreboard date, YYYYMMDD. Defaults to local today window.")
    parser.add_argument("--days-back", type=int, default=1)
    parser.add_argument("--days-forward", type=int, default=2)
    parser.add_argument("--allow-any", action="store_true", help="Use the first NBA event if no Finals event is found.")
    parser.add_argument("--board", default="Test", help="NewSMTH board. Use Basketball only after Test succeeds.")
    parser.add_argument("--no-player-stats", action="store_true", help="Do not include player boxscore stats.")
    parser.add_argument("--max-players", type=int, default=8, help="Maximum players to show per team.")
    parser.add_argument("--post", action="store_true", help="Actually post to NewSMTH. Without this, dry-run only.")
    parser.add_argument("--username", default=os.getenv("NEWSMTH_USER", ""))
    parser.add_argument("--password", default=os.getenv("NEWSMTH_PASS", ""))
    parser.add_argument("--yes", action="store_true", help="Skip the final local confirmation prompt.")
    parser.add_argument("--state-file", default="nba_finals_score_state.json")
    args = parser.parse_args()

    today = dt.date.today()
    dates = [args.date] if args.date else date_window(today, args.days_back, args.days_forward)
    if args.source == "yahoo":
        game = pick_yahoo_event(dates, args.allow_any)
        if args.format == "ptt":
            subject, body, state_key = build_yahoo_ptt_post(game, args.max_players, not args.no_player_stats)
        else:
            subject, body, state_key = build_yahoo_post(game, args.max_players, not args.no_player_stats)
        event_name = game.get("url")
        date_text = str(game.get("source_date"))
    else:
        date_text, event = pick_event(dates, args.allow_any)
        subject, body, state_key = build_post(event)
        event_name = event.get("name")

        if not args.no_player_stats:
            try:
                summary = fetch_summary(str(event.get("id") or ""))
                sections = boxscore_sections(summary, args.max_players)
                if sections:
                    body = "\n\n".join([body, *sections])
                else:
                    body = "\n\n".join([body, "Player stats: not available yet."])
            except RuntimeError as exc:
                body = "\n\n".join([body, f"Player stats: unavailable ({exc})"])

    state_path = Path(args.state_file)
    posted_keys = load_state(state_path)
    already_posted = state_key in posted_keys

    print(f"scoreboard date: {date_text}")
    print(f"source: {args.source}")
    print(f"event: {event_name}")
    print(f"state key: {state_key}")
    print()
    print("=== subject ===")
    print(subject)
    print()
    print("=== body ===")
    print(body)
    print()

    if already_posted:
        print("This exact score state was already posted according to the local state file.")
        return 0

    if not args.post:
        print("Dry run only. Add --post to publish.")
        return 0

    username = args.username or input("NewSMTH username: ")
    password = args.password or getpass.getpass("NewSMTH password: ")
    if not args.yes:
        answer = input(f"Type PUBLISH {args.board} to post this update: ")
        if answer.strip() != f"PUBLISH {args.board}":
            print("Canceled before posting.")
            return 1

    session = requests.Session()
    login_result = login_newsmth(session, username, password)
    if str(login_result.get("ajax_code")) != "0005":
        print("Login did not report success:")
        print(json.dumps(login_result, ensure_ascii=False, indent=2))
        return 1

    post_result = post_newsmth(session, args.board, subject, body)
    print("=== post result ===")
    print(json.dumps(post_result, ensure_ascii=False, indent=2))

    if str(post_result.get("ajax_code")) == "0005" or post_result.get("ajax_st") == 1:
        posted_keys.add(state_key)
        save_state(state_path, posted_keys)
        print(f"Saved posted state to {state_path}")
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())
