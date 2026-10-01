"""
bot_features_scraper.py
========================
Fetches robot DESIGN features (weapon type, drivetrain/wheels, weight, height,
length, width, motors, battery, armor, etc.) from the BattleBots Fandom wiki's
infobox on each robot's individual page.

This is a different scraping pattern than battlebots_scraper.py:
  - battlebots_scraper.py  -> 1 request per SEASON (fast, ~7 requests total)
  - bot_features_scraper.py -> 1 request per ROBOT (slower, 100-300+ requests)

Because of that, this script:
  - reads the list of robot names to look up from battlebots_2015_2025.xlsx
    (the "Career_Totals" sheet produced by battlebots_scraper.py)
  - caches every successful fetch to bot_features_cache.json, so if you stop
    and re-run the script, it won't re-download robots it already has
  - sleeps briefly between requests to be polite to Fandom's servers

Fandom wikis run on the same MediaWiki software as Wikipedia and render a
"portable infobox" (the info box on the top-right of a bot's page) as a
predictable HTML structure: <div class="portable-infobox"> containing
label/value pairs. This script parses that structure directly rather than
guessing at column names, which makes it resilient even though different
robots' infoboxes don't all have identical fields.

Usage:
    python bot_features_scraper.py
    (run battlebots_scraper.py first so there's a robot list to look up)

Output:
    Adds/overwrites a "Bot_Features" sheet in battlebots_2015_2025.xlsx
"""

import json
import os
import re
import sys
import time
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

WIKI = "https://battlebots.fandom.com"
API = f"{WIKI}/api.php"
HEADERS = {
    "User-Agent": "BattleBotsStatsResearch/1.0 (student research project; contact: youremail@example.com)"
}

WORKBOOK = "battlebots_2015_2025.xlsx"
CACHE_FILE = "bot_features_cache.json"

# ---------------------------------------------------------------------------
# Normalize the many different label wordings wikis use into consistent columns.
# Add to this dict as you notice new label variants in the raw data.
# ---------------------------------------------------------------------------
LABEL_MAP = {
    "weapon": "Weapon_Type", "weapon type": "Weapon_Type", "weapon(s)": "Weapon_Type",
    "weapon type(s)": "Weapon_Type",
    "weight": "Weight", "weight class": "Weight_Class",
    "height": "Height", "length": "Length", "width": "Width",
    "dimensions": "Dimensions",
    "drivetrain": "Drive_Type", "drive": "Drive_Type", "drive type": "Drive_Type",
    "locomotion": "Drive_Type", "chassis type": "Drive_Type",
    "speed": "Speed", "top speed": "Speed",
    "armor": "Armor", "armour": "Armor",
    "weapon power": "Weapon_Motor", "weapon motor": "Weapon_Motor",
    "drive power": "Drive_Motor", "drive motor": "Drive_Motor",
    "battery": "Battery", "batteries": "Battery",
    "team": "Team_Name", "builder": "Builder", "builder(s)": "Builder",
    "hometown": "Hometown", "nickname(s)": "Nickname",
}

WORD_TO_NUM = {
    "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8,
}

# Keyword -> canonical wheel-POSITION tag. Checked against the same Drive_Type
# text that Wheel_Count is extracted from. A robot's text can match more than
# one (e.g. "four internal wheels, corner-mounted" -> Internal + Corner-mounted),
# so this returns a list, not a single value.
WHEEL_POSITION_KEYWORDS = [
    (r"\binternal\b", "Internal"),
    (r"\bexternal\b", "External"),
    (r"\boutrigger", "Outrigger"),
    (r"\bcorner", "Corner-mounted"),
    (r"\bfront[\s-]?wheel", "Front-mounted"),
    (r"\brear[\s-]?wheel", "Rear-mounted"),
    (r"\bomni", "Omnidirectional"),
    (r"\bmecanum", "Omnidirectional"),
    (r"\bwalk(er|ing)\b|\bleg(ged|s)?\b", "Walker/Legged"),
    (r"\btrack(ed|s)?\b", "Tracked"),
    (r"\bhub[\s-]?motor", "Hub-motor"),
]



def get_robot_list(workbook=WORKBOOK):
    """Pull unique robot names to look up from the Career_Totals sheet."""
    if not os.path.exists(workbook):
        sys.exit(
            f"\nERROR: '{workbook}' does not exist yet.\n"
            f"This script needs the robot list that battlebots_scraper.py produces.\n"
            f"Run this first:\n\n    python battlebots_scraper.py\n\n"
            f"...and confirm it prints 'Done. Workbook saved to: {workbook}' before "
            f"running bot_features_scraper.py.\n"
        )
    try:
        df = pd.read_excel(workbook, sheet_name="Career_Totals")
    except ValueError as e:
        sys.exit(
            f"\nERROR: '{workbook}' exists but has no 'Career_Totals' sheet ({e}).\n"
            f"This usually means battlebots_scraper.py exited early/failed. Re-run it "
            f"and check its console output for 'FAILED' lines before continuing.\n"
        )
    names = sorted(df["Robot"].dropna().unique().tolist())
    if not names:
        sys.exit(
            f"\nERROR: 'Career_Totals' sheet in '{workbook}' has zero robots in it.\n"
            f"battlebots_scraper.py ran but collected no usable rows -- re-run it and "
            f"check its console output for details.\n"
        )
    return names


def load_cache():
    if Path(CACHE_FILE).exists():
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_cache(cache):
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, indent=2)


def fetch_infobox_html(robot_name: str):
    """Ask the Fandom API to render the page, return raw HTML or None if missing."""
    params = {
        "action": "parse",
        "page": robot_name,
        "prop": "text",
        "format": "json",
        "formatversion": "2",
    }
    resp = requests.get(API, params=params, headers=HEADERS, timeout=30)
    if resp.status_code != 200:
        return None
    data = resp.json()
    if "error" in data:
        return None
    return data["parse"]["text"]


def parse_portable_infobox(html: str) -> dict:
    """Extract every label/value pair from a Fandom 'portable infobox' block."""
    soup = BeautifulSoup(html, "lxml")
    box = soup.find("aside", class_=re.compile("portable-infobox"))
    if box is None:
        return {}

    fields = {}
    for item in box.find_all("div", class_=re.compile("pi-item")):
        label_el = item.find(class_=re.compile("pi-data-label"))
        value_el = item.find(class_=re.compile("pi-data-value"))
        if label_el and value_el:
            label = label_el.get_text(strip=True)
            value = value_el.get_text(separator=" ", strip=True)
            if label and value:
                fields[label] = value
    return fields


def normalize_fields(raw: dict) -> dict:
    """Map raw infobox labels onto a consistent schema; keep leftovers as Extra_*."""
    out = {}
    for label, value in raw.items():
        key = LABEL_MAP.get(label.strip().lower())
        if key:
            # Don't silently overwrite if the same canonical field appears twice
            if key in out:
                out[key] = out[key] + " / " + value
            else:
                out[key] = value
        else:
            out[f"Extra_{label}"] = value

    # Try to pull a wheel count out of whatever drive-type text we found
    drive_text = out.get("Drive_Type", "") or ""
    wheel_count = None
    match = re.search(r"(\b[a-z]+\b|\d+)[\s-]*wheel", drive_text.lower())
    if match:
        token = match.group(1)
        wheel_count = WORD_TO_NUM.get(token, None)
        if wheel_count is None and token.isdigit():
            wheel_count = int(token)
    out["Wheel_Count"] = wheel_count

    # Same idea for wheel POSITION/configuration -- e.g. "two internal wheels"
    # gives Wheel_Count=2, Wheel_Position="Internal". No match just means the
    # wiki's text didn't say (very common) -- left blank rather than guessed.
    positions_found = []
    for pattern, tag in WHEEL_POSITION_KEYWORDS:
        if re.search(pattern, drive_text.lower()):
            positions_found.append(tag)
    out["Wheel_Position"] = " / ".join(positions_found) if positions_found else None

    return out


def build_features_dataframe(robot_names):
    cache = load_cache()
    rows = []

    for name in robot_names:
        if name in cache:
            print(f"  (cached) {name}")
            record = cache[name]
        else:
            print(f"Fetching features for: {name}")
            html = fetch_infobox_html(name)
            if html is None:
                print(f"  WARNING: no Fandom page found for '{name}' — skipping")
                cache[name] = {"__not_found__": True}
                save_cache(cache)
                time.sleep(0.5)
                continue
            raw = parse_portable_infobox(html)
            if not raw:
                print(f"  WARNING: no infobox found on page for '{name}'")
            record = normalize_fields(raw)
            cache[name] = record
            save_cache(cache)
            time.sleep(0.5)  # be polite

        if record.get("__not_found__"):
            continue
        record = dict(record)  # copy
        record["Robot"] = name
        rows.append(record)

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    # Put the most analysis-relevant columns first, extras trail at the end
    priority = ["Robot", "Weapon_Type", "Drive_Type", "Wheel_Count", "Weight",
                "Weight_Class", "Height", "Length", "Width", "Speed", "Armor",
                "Weapon_Motor", "Drive_Motor", "Battery"]
    ordered = [c for c in priority if c in df.columns]
    rest = [c for c in df.columns if c not in ordered]
    return df[ordered + rest]


def append_to_workbook(features_df, workbook=WORKBOOK, sheet_name="Bot_Features"):
    from openpyxl import load_workbook

    wb = load_workbook(workbook)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
        wb.save(workbook)

    with pd.ExcelWriter(workbook, engine="openpyxl", mode="a") as writer:
        features_df.to_excel(writer, sheet_name=sheet_name, index=False)

    import excel_style
    excel_style.style_workbook(workbook)
    print(f"\n'{sheet_name}' sheet written to {workbook} ({len(features_df)} robots).")


if __name__ == "__main__":
    names = get_robot_list()
    print(f"Found {len(names)} unique robots in Career_Totals. Fetching design features...\n")
    features_df = build_features_dataframe(names)
    if features_df.empty:
        print("No feature data collected.")
    else:
        append_to_workbook(features_df)