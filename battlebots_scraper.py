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
    battlebots_2015_2025.xlsx  (in the same folder)

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
# 3. PARSE: find the Contestants table among all tables on the page
# ---------------------------------------------------------------------------
def find_contestants_table(html: str, title: str) -> pd.DataFrame:
    try:
        # IMPORTANT: pass HTML through StringIO, not as a raw string. Newer
        # pandas/lxml versions can otherwise mistake certain HTML strings for
        # a file path and raise a confusing FileNotFoundError instead of
        # actually parsing the markup.
        tables = pd.read_html(StringIO(html))
    except ImportError as e:
        raise ImportError(
            "pandas.read_html needs 'lxml' installed to parse HTML tables. "
            "Run: pip install -r requirements.txt  (then try again)"
        ) from e
    except ValueError as e:
        raise ValueError(f"pandas found no HTML tables at all on '{title}': {e}")

    ROBOT_HINTS = ("robot", "bot name", "name")
    RECORD_HINTS = ("record", "fight", "w-l", "w/l")

    candidates = []
    for i, df in enumerate(tables):
        cols = [str(c).strip().lower() for c in df.columns]
        has_robot = any(any(h in c for h in ROBOT_HINTS) for c in cols)
        has_record = any(any(h in c for h in RECORD_HINTS) for c in cols)
        candidates.append((i, cols, has_robot, has_record))
        if has_robot and has_record:
            return df

    # Nothing matched both hints -- print full diagnostics so this is fixable
    # instead of a silent skip.
    print(f"    Could not auto-detect the Contestants table on '{title}'. "
          f"Found {len(tables)} table(s) on the page:")
    for i, cols, has_robot, has_record in candidates:
        print(f"      Table {i}: columns = {cols}  "
              f"(robot-like column: {has_robot}, record-like column: {has_record})")

    raise ValueError(
        f"No table with both a Robot-like column and a Record-like column was found on '{title}'. "
        f"See the printed column list above -- if one of those tables is clearly the right one "
        f"but named differently, update ROBOT_HINTS/RECORD_HINTS in find_contestants_table()."
    )


def clean_record(record):
    """Turn '3-1' or '1-2-1' (various dash characters) into (wins, losses, ties)."""
    if pd.isna(record):
        return None, None, None
    normalized = re.sub(r"[\u2012\u2013\u2014\u2212]", "-", str(record)).strip()
    parts = re.findall(r"\d+", normalized)
    if len(parts) == 2:
        wins, losses = map(int, parts)
        return wins, losses, 0
    elif len(parts) == 3:
        wins, losses, ties = map(int, parts)
        return wins, losses, ties
    return None, None, None


def normalize_columns(df: pd.DataFrame, title: str) -> pd.DataFrame:
    """Map whatever the real column names are to a consistent schema."""
    # Flatten MultiIndex columns (can happen with rowspan'd wiki table headers)
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [" ".join(str(x) for x in tup if "Unnamed" not in str(x)).strip()
                      for tup in df.columns]

    rename_map = {}
    for c in df.columns:
        cl = str(c).strip().lower()
        if "robot" in cl or cl == "name" or cl == "bot name":
            rename_map[c] = "Robot"
        elif "builder" in cl:
            rename_map[c] = "Builder"
        elif "hometown" in cl or "location" in cl:
            rename_map[c] = "Hometown"
        elif "record" in cl or "fight" in cl or cl in ("w-l", "w/l"):
            rename_map[c] = "Fight Record"
    df = df.rename(columns=rename_map)

    if "Robot" not in df.columns:
        raise ValueError(
            f"After renaming, '{title}' still has no 'Robot' column. "
            f"Original columns were: {list(df.columns)}"
        )
    if "Fight Record" not in df.columns:
        raise ValueError(
            f"After renaming, '{title}' still has no 'Fight Record' column. "
            f"Original columns were: {list(df.columns)}"
        )

    keep = [c for c in ["Robot", "Builder", "Hometown", "Fight Record"] if c in df.columns]
    return df[keep]


# ---------------------------------------------------------------------------
# 4. MAIN PIPELINE
# ---------------------------------------------------------------------------
def build_dataset():
    all_rows = []
    per_season_frames = {}
    failures = []

    for title, label, year in SEASONS:
        print(f"Fetching {label} ...")
        try:
            html = fetch_page_html(title)
            raw_df = find_contestants_table(html, title)
            df = normalize_columns(raw_df, title)
        except Exception as e:
            print(f"  FAILED: {label} -> {e}\n")
            failures.append((label, str(e)))
            continue

        df["Robot"] = df["Robot"].astype(str).str.replace(r"\[.*?\]", "", regex=True).str.strip()
        df["Robot"] = df["Robot"].str.replace(r"\*+$", "", regex=True).str.strip()

        wins, losses, ties = [], [], []
        for rec in df.get("Fight Record", pd.Series([None] * len(df))):
            w, l, t = clean_record(rec)
            wins.append(w)
            losses.append(l)
            ties.append(t)

        df["Wins"] = wins
        df["Losses"] = losses
        df["Ties"] = ties
        df["Total Fights"] = df[["Wins", "Losses", "Ties"]].sum(axis=1, min_count=1)
        df["Win %"] = (df["Wins"] / df["Total Fights"]).round(3)
        df["Season"] = label
        df["Year"] = year

        per_season_frames[label] = df.copy()
        all_rows.append(df)
        print(f"  OK: {label} -> {len(df)} robots parsed")
        time.sleep(0.5)  # be polite to Wikipedia's servers

    print("\n" + "=" * 70)
    print(f"SUMMARY: {len(all_rows)}/{len(SEASONS)} seasons parsed successfully.")
    if failures:
        print("Seasons that FAILED:")
        for label, err in failures:
            print(f"  - {label}: {err}")
    print("=" * 70 + "\n")

    if not all_rows:
        print(
            "No data was collected at all, so no Excel file was written.\n"
            "Read the 'FAILED' messages above -- they tell you exactly what went wrong "
            "for each season (network error vs. table-not-found vs. missing column).\n"
            "Most common causes:\n"
            "  1. No internet access from this machine/environment right now.\n"
            "  2. A firewall/proxy blocking en.wikipedia.org.\n"
            "  3. Wikipedia changed a page's table structure (the printed column "
            "list above each failure shows you exactly what's on the page now).\n"
        )
        sys.exit(1)

    all_fights = pd.concat(all_rows, ignore_index=True)
    ordered_cols = ["Season", "Year", "Robot", "Builder", "Hometown",
                     "Wins", "Losses", "Ties", "Total Fights", "Win %"]
    all_fights = all_fights[[c for c in ordered_cols if c in all_fights.columns]]

    career = (
        all_fights.groupby("Robot", as_index=False)
        .agg(
            Seasons_Played=("Season", lambda s: ", ".join(sorted(set(s)))),
            Number_of_Seasons=("Season", "nunique"),
            Career_Wins=("Wins", "sum"),
            Career_Losses=("Losses", "sum"),
            Career_Ties=("Ties", "sum"),
        )
    )
    career["Career_Fights"] = career[["Career_Wins", "Career_Losses", "Career_Ties"]].sum(axis=1)
    career["Career_Win %"] = (career["Career_Wins"] / career["Career_Fights"]).round(3)
    career = career.sort_values("Career_Wins", ascending=False).reset_index(drop=True)

    return all_fights, career, per_season_frames


# ---------------------------------------------------------------------------
# 5. EXPORT TO EXCEL
# ---------------------------------------------------------------------------
def export_to_excel(all_fights, career, per_season_frames, path=OUTPUT_FILE):
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        all_fights.to_excel(writer, sheet_name="All_Fights_Raw", index=False)
        career.to_excel(writer, sheet_name="Career_Totals", index=False)
        for label, df in per_season_frames.items():
            sheet_name = label[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)

    _autofit_and_style(path)
    print(f"Done. Workbook saved to: {path}")


def _autofit_and_style(path):
    from openpyxl import load_workbook
    from openpyxl.styles import Font

    wb = load_workbook(path)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for col_cells in ws.columns:
            length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
            col_letter = col_cells[0].column_letter
            ws.column_dimensions[col_letter].width = min(max(length + 2, 10), 45)
    wb.save(path)


if __name__ == "__main__":
    fights_df, career_df, season_dfs = build_dataset()
    export_to_excel(fights_df, career_df, season_dfs)
