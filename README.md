# BattleBots Stats Scraper (2015–present)

Automates pulling win/loss records for every BattleBots robot across the modern era
(2015 reboot onward) into a single Excel workbook, instead of copying data one bot
page at a time.

## Why Wikipedia as the source

Individual bot pages (on BattleBots.com or the BattleBots Fandom wiki) only show
one robot's history at a time. Wikipedia's season pages (e.g. `BattleBots season 12`)
each contain a single **Contestants table** with every robot's builder, hometown,
and season fight record (e.g. `3-1`) already in one place — which is what makes
this scrapable in bulk.

## Setup

```bash
git clone <this-repo>
cd battlebots-scraper
pip install -r requirements.txt
```

Before running, open `battlebots_scraper.py` and put a real contact email in the
`HEADERS["User-Agent"]` string — Wikipedia's API etiquette asks for this so they
can reach you if your script misbehaves, and some proxies rate-limit or block
generic/blank user agents.

## Run the full pipeline

Run these three scripts **in order**. Each one adds sheets to the same workbook,
so later scripts depend on the ones before them.

```bash
python battlebots_scraper.py        # 1. win/loss records, by season
python bot_features_scraper.py      # 2. design specs, per robot
python analyze_correlations.py      # 3. win-rate vs. feature correlation
```

This produces `battlebots_2015_2025.xlsx` with:

| Sheet                  | Contents                                                                 | From script |
|------------------------|---------------------------------------------------------------------------|-------------|
| `All_Fights_Raw`       | One row per robot per season: Wins, Losses, Ties, Win %                   | 1 |
| `Career_Totals`        | One row per robot, aggregated across every season it competed in         | 1 |
| `Season N (year)`      | One sheet per individual season, for quick browsing                      | 1 |
| `Bot_Features`         | One row per robot: Weapon_Type, Drive_Type, Wheel_Count, Weight, Height, Length, Width, Speed, Armor, motors, battery, plus any extra infobox fields the wiki happened to have | 2 |
| `Feature_Correlations` | For every feature value (e.g. "Horizontal Bar Spinner"), how its average win % compares to the field average ("Lift") | 3 |

### Why steps 2 and 3 are separate scripts, not folded into step 1

Step 1 fetches ~7 Wikipedia pages (one per season) — fast and light.
Step 2 fetches one Fandom page **per robot** (100–300+ requests depending on
how many years you cover) — much slower, and it's a different site with a
different HTML structure (a "portable infobox," not a plain table). Splitting
them means you can re-run step 3's analysis instantly without re-scraping,
and step 2 caches every robot it successfully fetches to
`bot_features_cache.json` so interrupting/restarting doesn't cost you progress.

### How the correlation numbers work (put this in your report)

For each feature value, `analyze_correlations.py` computes:

```
lift = (average career win% of robots with that feature) / (average career win% of ALL robots)
```

`lift > 1` means robots with that feature won more than the field average;
`lift < 1` means they won less. This is a simple ratio-of-averages, **not**
a formal statistical test — it doesn't control for era, weight class, or
number of fights, and a feature used by only 2-3 robots can produce a
misleadingly extreme lift score just from small-sample noise. The
`Small_Sample_Warning` column flags any group with fewer than 5 robots;
the console output when you run the script already filters those out of
the "top 10" lists, but they're still in the Excel sheet in case you want
to eyeball them.

## Extending it

- **Add a newer season:** once Wikipedia publishes a finished "Contestants" table
  for a season not yet in the script, add a tuple to the `SEASONS` list at the top
  of `battlebots_scraper.py`.
- **Add more feature columns to the correlation analysis:** if `Bot_Features`
  ends up with useful `Extra_*` columns (some robots' infoboxes have unique
  fields others don't), add the column name to `FEATURE_COLUMNS` at the top of
  `analyze_correlations.py`.
- **Robustness:** the season parser looks for a table containing "Robot" +
  "Record"/"Fight" columns rather than hardcoding table position, and the
  feature parser reads the Fandom infobox's actual label/value HTML structure
  rather than guessing field order — both should survive minor page edits.
  If a page changes significantly, the relevant script prints a warning and
  skips that item rather than crashing the whole run.

## Notes on the design-feature data specifically

- Not every robot's Fandom page has a complete infobox — some older or
  short-lived robots have sparse pages. Missing fields just show up as empty
  cells; the script doesn't invent data.
- Some infoboxes list more than one weapon type over a robot's history (e.g.
  a bot that switched from a spinner to a flipper between seasons) as a
  single text field. `analyze_correlations.py` splits on slashes/commas/"and"
  so each weapon type still gets counted, but this means one robot can
  contribute to more than one feature group — worth noting as a limitation
  if you cite exact counts in your report.
- `Wheel_Count` is extracted with a regex looking for phrases like "two-wheel
  drive" in the drivetrain text — it won't catch every phrasing a wiki editor
  used, so spot-check a sample against the raw `Drive_Type` column before
  trusting it heavily.

## Notes on data quality (worth mentioning in your report)

- "Fight Record" reflects **regular-season/qualifying fights only** for most
  seasons — postseason bracket results are usually in a separate table on the
  same page, not folded into this column. Check a season's Wikipedia page
  manually if you need bracket-specific wins/losses (e.g. who won the Giant Nut).
- A few robots changed names or teams between seasons (e.g. rebuilds); the
  `Career_Totals` sheet groups strictly by exact robot name string, so minor
  naming inconsistencies on Wikipedia could split one robot's history into two
  rows. Worth a manual sanity check for bots you're specifically writing about.
- This is intended as a research aid for a class assignment, not a definitive
  competition database — always double check key stats you cite against the
  primary source (BattleBots.com or the original broadcast) before including
  them in your final report.

## License / attribution

Data pulled from Wikipedia (CC BY-SA). If you publish this workbook or derived
figures, credit Wikipedia contributors as the data source.
