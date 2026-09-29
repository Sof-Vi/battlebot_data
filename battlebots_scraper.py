"""
battlebots_scraper.py
======================
Pulls per-season "Contestants" tables (Robot / Builder / Hometown / Fight Record)
for every modern-era BattleBots season (2015-present) straight from Wikipedia,
cleans them up, and writes one tidy Excel workbook with:

  - Sheet "All_Fights_Raw"       -> one row per robot per season
  - Sheet "Career_Totals"        -> one row per robot, career W/L totals + seasons played
  - Sheet "<Season Name>"        -> one sheet per season, for quick per-season browsing

Usage:
    pip install -r requirements.txt
    python battlebots_scraper.py

Output:
    battlebots_2018_2022.xlsx  (in the same folder)

IF THIS SCRIPT PRODUCES NO DATA / EXITS WITH AN ERROR:
    This version prints a diagnostic report for every season it could not
    parse -- read that output first. It will tell you exactly one of:
      (a) the network request itself failed (no internet / blocked / rate-limited)
      (b) the page loaded fine, but no table with "Robot" + a record-like column
          was found -- in which case it also PRINTS every table's column names
          it did find on that page, so you can see what actually changed and
          fix the matching rule in normalize_columns()/find_contestants_table()
          yourself, or paste that printed output back for a fix.
"""

import re
import time
import sys
from io import StringIO

import requests
import pandas as pd

# ---------------------------------------------------------------------------
# 1. CONFIG: which seasons to pull, and what real-world year each corresponds to
# ---------------------------------------------------------------------------
SEASONS = [
    ("BattleBots season 6",  "Season 6 (2015)",  2015),
    ("BattleBots season 7",  "Season 7 (2016)",  2016),
    ("BattleBots season 8",  "Season 8 (2018)",  2018),
    ("BattleBots season 9",  "Season 9 (2019)",  2019),
    ("BattleBots season 10", "Season 10 (2020)", 2020),
    ("BattleBots season 11", "Season 11 (2021)", 2021),
    ("BattleBots season 12", "Season 12 (2022)", 2022),
    # Add newer seasons here once Wikipedia publishes a finished Contestants table, e.g.:
    # ("BattleBots season 13", "Season 13 (2024)", 2024),
]

WIKI_API = "https://en.wikipedia.org/w/api.php"
HEADERS = {
    "User-Agent": "BattleBotsStatsResearch/1.0 (student research project; contact: youremail@example.com)"
}
OUTPUT_FILE = "battlebots_2015_2025.xlsx"
MAX_RETRIES = 3


# ---------------------------------------------------------------------------
# 2. FETCH: get the rendered HTML of a Wikipedia page via the API, with retries
# ---------------------------------------------------------------------------
def fetch_page_html(title: str) -> str:
    params = {
        "action": "parse",
        "page": title,
        "prop": "text",
        "format": "json",
        "formatversion": "2",
    }
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(WIKI_API, params=params, headers=HEADERS, timeout=30)
        except requests.exceptions.RequestException as e:
            last_error = f"network error ({e.__class__.__name__}): {e}"
            print(f"    attempt {attempt}/{MAX_RETRIES} failed: {last_error}")
            time.sleep(1.5 * attempt)
            continue

        if resp.status_code == 429:
            last_error = "rate-limited (HTTP 429) by Wikipedia"
            print(f"    attempt {attempt}/{MAX_RETRIES}: {last_error}, backing off...")
            time.sleep(3 * attempt)
            continue
        if resp.status_code != 200:
            last_error = f"HTTP {resp.status_code}: {resp.text[:200]}"
            print(f"    attempt {attempt}/{MAX_RETRIES} failed: {last_error}")
            time.sleep(1.5 * attempt)
            continue

        try:
            data = resp.json()
        except ValueError:
            last_error = f"response was not valid JSON (first 200 chars): {resp.text[:200]}"
            print(f"    attempt {attempt}/{MAX_RETRIES} failed: {last_error}")
            continue

        if "error" in data:
            # This usually means the page title doesn't exist / was renamed.
            raise ValueError(
                f"Wikipedia says page '{title}' doesn't exist or API rejected it: {data['error']}. "
                f"Double check the exact title at https://en.wikipedia.org/wiki/{title.replace(' ', '_')}"
            )

        return data["parse"]["text"]

    raise ConnectionError(
        f"Could not fetch '{title}' after {MAX_RETRIES} attempts. Last error: {last_error}. "
        f"This usually means either you have no internet connection right now, or your "
        f"network/firewall is blocking requests to en.wikipedia.org."
    )


# ---------------------------------------------------------------------------
# 3. PARSE
# ---------------------------------------------------------------------------
# Wikipedia season pages are NOT uniform:
#   Seasons 8-12 : Contestants table has a "Fight Record" column (e.g. 3-1)
#                  and the table is SPLIT into several pieces on the page.
#   Seasons 6-7  : Contestants table has NO "Fight Record" column (it has
#                  Weapon / Elim. in instead), but the page has full fight-by-
#                  fight result tables (Winner / Loser / Method / Time).
# So we (a) collect ALL contestant table pieces, and (b) when there's no
# Fight Record column, we compute every robot's record from the fight tables.

def _flatten_columns(df):
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join(str(x) for x in tup if "Unnamed" not in str(x)).strip()
                      for tup in df.columns]
    else:
        df.columns = [str(c).strip() for c in df.columns]
    return df


def robot_key(name) -> str:
    """Matching key so 'Death Roll' == 'DeathRoll', 'The Ringmaster' == 'Ringmaster'."""
    s = re.sub(r"\[.*?\]", "", str(name)).lower().strip()
    s = re.sub(r"^the\s+", "", s)
    return re.sub(r"[^a-z0-9]", "", s)


def clean_robot_name(name) -> str:
    s = re.sub(r"\[.*?\]", "", str(name)).strip()
    return re.sub(r"[\*\u2020\u2021]+$", "", s).strip()


def read_all_tables(html: str, title: str):
    try:
        # StringIO avoids pandas mistaking the HTML string for a file path
        tables = pd.read_html(StringIO(html))
    except ImportError as e:
        raise ImportError("pandas.read_html needs 'lxml'. Run: pip install -r requirements.txt") from e
    except ValueError as e:
        raise ValueError(f"pandas found no HTML tables at all on '{title}': {e}")
    return [_flatten_columns(t) for t in tables]


def find_contestant_tables(tables, title):
    """Return ALL table pieces that list robots (they are often split in two)."""
    pieces, diagnostics = [], []
    for i, df in enumerate(tables):
        cols = [c.lower() for c in df.columns]
        has_robot = any("robot" in c for c in cols)
        has_result = any(("record" in c) or ("elim" in c) for c in cols)
        diagnostics.append((i, cols))
        if has_robot and has_result:
            pieces.append(df)
    if not pieces:
        print(f"    No contestant table found on '{title}'. Tables on page:")
        for i, cols in diagnostics:
            print(f"      Table {i}: {cols}")
        raise ValueError(f"No table with a Robot column and a Record/Elim. column on '{title}'.")
    return pieces


def normalize_contestants(pieces, title) -> pd.DataFrame:
    frames = []
    for df in pieces:
        rename = {}
        for c in df.columns:
            cl = c.lower()
            if "robot" in cl:                 rename[c] = "Robot"
            elif "weapon" in cl:              rename[c] = "Weapon"
            elif "builder" in cl:             rename[c] = "Builder"
            elif "hometown" in cl or "location" in cl: rename[c] = "Hometown"
            elif "record" in cl:              rename[c] = "Fight Record"
            elif "elim" in cl:                rename[c] = "Eliminated In"
        df = df.rename(columns=rename)
        df = df.loc[:, ~df.columns.duplicated()]
        keep = [c for c in ["Robot", "Weapon", "Builder", "Hometown",
                            "Fight Record", "Eliminated In"] if c in df.columns]
        frames.append(df[keep])
    out = pd.concat(frames, ignore_index=True)
    # drop repeated header rows / junk rows that ended up inside the data
    out = out[out["Robot"].notna()]
    out = out[out["Robot"].astype(str).str.strip().str.lower() != "robot"]
    out["Robot"] = out["Robot"].map(clean_robot_name)
    out = out[out["Robot"] != ""].drop_duplicates(subset="Robot").reset_index(drop=True)
    return out


def find_fight_tables(tables) -> pd.DataFrame:
    """
    Fight-by-fight tables: columns like Episode | Battle | Winner | Loser | Method | Time.
    Rumble tables (3+ robots) use 'Losers' and are skipped on purpose.
    """
    rows = []
    for df in tables:
        cols = {c.lower(): c for c in df.columns}
        win_col = cols.get("winner") or cols.get("champion")
        lose_col = cols.get("loser")
        if not win_col or not lose_col:
            continue
        for _, r in df.iterrows():
            w, l = r[win_col], r[lose_col]
            if pd.isna(w) or pd.isna(l):
                continue
            method = str(r[cols["method"]]) if "method" in cols else ""
            method = re.sub(r"\[.*?\]|\^|\{.*?\}", "", method).strip()
            rows.append({
                "Episode": r[cols["episode"]] if "episode" in cols else None,
                "Winner": clean_robot_name(w),
                "Loser": clean_robot_name(l),
                "Method": method,
                "Time": r[cols["time"]] if "time" in cols else None,
            })
    return pd.DataFrame(rows)


def records_from_fights(contestants: pd.DataFrame, fights: pd.DataFrame) -> pd.DataFrame:
    wins, losses = {}, {}
    for _, f in fights.iterrows():
        wins[robot_key(f["Winner"])] = wins.get(robot_key(f["Winner"]), 0) + 1
        losses[robot_key(f["Loser"])] = losses.get(robot_key(f["Loser"]), 0) + 1
    keys = contestants["Robot"].map(robot_key)
    contestants = contestants.copy()
    contestants["Wins"] = keys.map(lambda k: wins.get(k, 0))
    contestants["Losses"] = keys.map(lambda k: losses.get(k, 0))
    contestants["Ties"] = 0
    known = set(keys)
    unmatched = (set(wins) | set(losses)) - known
    if unmatched:
        print(f"    note: {len(unmatched)} name(s) in fight tables not in contestants table "
              f"(spelling differences?): {sorted(unmatched)[:8]}")
    return contestants


def clean_record(record):
    """Turn '3-1' or '1-2-1' (various dash characters) into (wins, losses, ties)."""
    if pd.isna(record):
        return None, None, None
    normalized = re.sub(r"[\u2012\u2013\u2014\u2212]", "-", str(record)).strip()
    parts = re.findall(r"\d+", normalized)
    if len(parts) == 2:
        return int(parts[0]), int(parts[1]), 0
    if len(parts) == 3:
        return int(parts[0]), int(parts[1]), int(parts[2])
    return None, None, None


# ---------------------------------------------------------------------------
# 4. MAIN PIPELINE
# ---------------------------------------------------------------------------
def build_dataset():
    all_rows, fight_logs, per_season_frames, failures = [], [], {}, []

    for title, label, year in SEASONS:
        print(f"Fetching {label} ...")
        try:
            html = fetch_page_html(title)
            tables = read_all_tables(html, title)
            df = normalize_contestants(find_contestant_tables(tables, title), title)
            fights = find_fight_tables(tables)

            if "Fight Record" in df.columns:
                recs = [clean_record(x) for x in df["Fight Record"]]
                df["Wins"] = [r[0] for r in recs]
                df["Losses"] = [r[1] for r in recs]
                df["Ties"] = [r[2] for r in recs]
                df["Record Source"] = "Contestants table (Fight Record)"
            elif not fights.empty:
                df = records_from_fights(df, fights)
                df["Record Source"] = "Computed from fight-by-fight tables"
            else:
                raise ValueError("No Fight Record column AND no fight-by-fight tables to compute one from.")
        except Exception as e:
            print(f"  FAILED: {label} -> {e}\n")
            failures.append((label, str(e)))
            continue

        df["Total Fights"] = df[["Wins", "Losses", "Ties"]].sum(axis=1, min_count=1)
        df["Win %"] = (df["Wins"] / df["Total Fights"].where(df["Total Fights"] > 0)).round(3)
        df["Season"], df["Year"] = label, year

        if not fights.empty:
            fl = fights.copy()
            fl.insert(0, "Year", year)
            fl.insert(0, "Season", label)
            fight_logs.append(fl)

        per_season_frames[label] = df.copy()
        all_rows.append(df)
        print(f"  OK: {label} -> {len(df)} robots ({df['Record Source'].iloc[0]})"
              + (f", {len(fights)} individual fights logged" if not fights.empty else ""))
        time.sleep(0.5)

    print("\n" + "=" * 70)
    print(f"SUMMARY: {len(all_rows)}/{len(SEASONS)} seasons parsed successfully.")
    for label, err in failures:
        print(f"  FAILED - {label}: {err}")
    print("=" * 70 + "\n")

    if not all_rows:
        print("No data was collected, so no Excel file was written. Read the FAILED lines above.")
        sys.exit(1)

    all_fights = pd.concat(all_rows, ignore_index=True)
    ordered = ["Season", "Year", "Robot", "Weapon", "Builder", "Hometown", "Eliminated In",
               "Wins", "Losses", "Ties", "Total Fights", "Win %", "Record Source"]
    all_fights = all_fights[[c for c in ordered if c in all_fights.columns]]
    all_fights["Robot_Key"] = all_fights["Robot"].map(robot_key)

    def first_non_null(s):
        s = s.dropna()
        return s.iloc[0] if len(s) else None

    career = (
        all_fights.groupby("Robot_Key", as_index=False)
        .agg(
            Robot=("Robot", "last"),      # most recent spelling
            Seasons_Played=("Season", lambda s: ", ".join(sorted(set(s)))),
            Number_of_Seasons=("Season", "nunique"),
            Career_Wins=("Wins", "sum"),
            Career_Losses=("Losses", "sum"),
            Career_Ties=("Ties", "sum"),
            Weapon_S6_S7=("Weapon", first_non_null) if "Weapon" in all_fights.columns else ("Season", "first"),
        )
    )
    if "Weapon" not in all_fights.columns:
        career = career.drop(columns=["Weapon_S6_S7"])
    career["Career_Fights"] = career[["Career_Wins", "Career_Losses", "Career_Ties"]].sum(axis=1)
    career["Career_Win %"] = (career["Career_Wins"] / career["Career_Fights"].where(career["Career_Fights"] > 0)).round(3)
    front = ["Robot", "Seasons_Played", "Number_of_Seasons", "Career_Wins", "Career_Losses",
             "Career_Ties", "Career_Fights", "Career_Win %"]
    career = career[front + [c for c in career.columns if c not in front + ["Robot_Key"]]]
    career = career.sort_values("Career_Wins", ascending=False).reset_index(drop=True)
    all_fights = all_fights.drop(columns=["Robot_Key"])

    fight_log = pd.concat(fight_logs, ignore_index=True) if fight_logs else pd.DataFrame()
    return all_fights, career, per_season_frames, fight_log


# ---------------------------------------------------------------------------
# 5. EXPORT TO EXCEL
# ---------------------------------------------------------------------------
def export_to_excel(all_fights, career, per_season_frames, fight_log=None, path=OUTPUT_FILE):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        all_fights.to_excel(writer, sheet_name="All_Fights_Raw", index=False)
        career.to_excel(writer, sheet_name="Career_Totals", index=False)
        if fight_log is not None and not fight_log.empty:
            fight_log.to_excel(writer, sheet_name="Fight_Log", index=False)
        for label, df in per_season_frames.items():
            sheet_name = label[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    import excel_style
    excel_style.style_workbook(path)
    print(f"Done. Workbook saved to: {path}")


if __name__ == "__main__":
    fights_df, career_df, season_dfs, fight_log_df = build_dataset()
    export_to_excel(fights_df, career_df, season_dfs, fight_log_df)