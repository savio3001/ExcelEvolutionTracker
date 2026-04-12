# Excel Evolution Tracker

Detect structural drift across monthly XLSB template files. Built to feed the results into a production data-scraping application's mapping update workflow.

The system extracts a structural fingerprint from each XLSB, compares consecutive versions, and reports changes that materially affect downstream consumers — sheet renames, block moves, label edits, named range additions, merged region changes, and so on. A learning classifier filters out value churn (monthly version numbers, dates, totals) so reports stay focused on changes that actually break scraping mappings.

## What gets tracked

- **Sheets** — added, removed, renamed (fuzzy + content cross-check), reordered
- **Blocks** — added, removed, moved, reshaped, primary label changed, internal labels changed
- **Named ranges** — added, removed (address shifts ignored by default)
- **Merged regions** — added, removed per sheet

Formula tracking is intentionally excluded.

## Install

```bash
uv sync
```

Creates a virtual environment and installs all dependencies from `pyproject.toml`. Windows + Excel are required for the XLSB → XLSX conversion step (uses `pywin32`). Everything downstream is cross-platform.

To include dev tools (pytest):

```bash
uv sync --extra dev
```

## Quick start

### One-time backfill over your existing files

```bash
uv run excel-evo-tracker batch --input ./xlsb_files/ --reset-db
```

Converts every XLSB, extracts snapshots, detects blocks, generates sequential diffs, persists to JSON + SQLite, writes Markdown + CSV reports for each diff, and produces a rollup index. The stability classifier learns from each diff as the batch progresses. For best results, run on at least 10–15 consecutive monthly files so the classifier has enough history to start filtering value noise.

### Monthly cadence — process one new file

```bash
uv run excel-evo-tracker incremental ./xlsb_files/template_2024_05.xlsb
```

Compares the new file against the most recent stored snapshot, generates the diff, updates the classifier, and refreshes the rollup.

### Ad-hoc compare without touching the DB

```bash
uv run excel-evo-tracker compare --old jan.xlsb --new feb.xlsb
```

Useful for inspecting two files side by side without affecting the persistent history.

## Commands

| Command | Purpose |
|---------|---------|
| `batch` | Process every XLSB in a directory; build full history |
| `incremental` | Process one new XLSB and diff against the latest stored |
| `compare` | Diff two XLSB files directly without touching the DB |
| `rollup` | Regenerate the cross-diff rollup index |
| `history` | Show recorded changes for a named element across all files |
| `review` | List all changes flagged `needs_review` |
| `stability` | Audit the learned cell stability classifier |
| `timeline` | Generate a timeline report for a single block on a single sheet |
| `blocks` | List blocks detected on a sheet from a stored snapshot |

Run `uv run excel-evo-tracker <command> --help` for details on any command.

## Inspecting your data

### Browse blocks on a sheet

```bash
# All blocks on a sheet (most recent snapshot)
uv run excel-evo-tracker blocks "Inputs"

# Filter by primary label substring (case-insensitive)
uv run excel-evo-tracker blocks "Inputs" --filter "rate"

# Filter by sheet region (positional range argument)
uv run excel-evo-tracker blocks "Inputs" "CS1:DA40"   # cell range
uv run excel-evo-tracker blocks "Inputs" "CS:DA"      # entire columns
uv run excel-evo-tracker blocks "Inputs" "1:40"       # entire rows

# Combine filters
uv run excel-evo-tracker blocks "Inputs" "CS:DA" --filter "rate"

# Read from a specific historical snapshot instead of the latest
uv run excel-evo-tracker blocks "Inputs" --file template_2024_03.xlsb
```

### Trace a single block across all months

```bash
# All recorded changes for "Version Number" on sheet "Header"
uv run excel-evo-tracker timeline Header "Version Number"

# Limit to the most recent N changes
uv run excel-evo-tracker timeline Header "Version Number" --limit 6

# Fuzzy block name match if you don't remember it exactly
uv run excel-evo-tracker timeline Header "version" --fuzzy

# Custom output path
uv run excel-evo-tracker timeline Header "Version Number" --out my_timeline.md
```

Produces a Markdown report and a sibling CSV in `reports/`.

### Search by element name across all categories

```bash
uv run excel-evo-tracker history "Net Revenue"
```

Walks every recorded change (sheets, blocks, named ranges, merges) where the element name matches.

### Triage changes flagged for human review

```bash
uv run excel-evo-tracker review
```

Lists every change the system wasn't fully confident about — low matching confidence or uncertain classifier verdict.

### Audit the learned classifier

```bash
# Dump every tracked sheet
uv run excel-evo-tracker stability

# One sheet
uv run excel-evo-tracker stability Header

# Multiple sheets — quote names with spaces
uv run excel-evo-tracker stability "Inputs" "Lookup Tables"

# Comma-separated alternative
uv run excel-evo-tracker stability --sheets "Header,Summary,Inputs"

# Fuzzy sheet name matching
uv run excel-evo-tracker stability "head" --fuzzy
```

Shows each tracked cell with its learned role (LABEL / VALUE / UNKNOWN), stability score, and observation count.

## What you get

For each diff, three files are produced under `reports/`:

- **`<old>__to__<new>.md`** — Markdown report with a Scraping Impact section at the top, then a full breakdown by category. Headers show the old/new file names.
- **`<old>__to__<new>.csv`** — one row per change. Open in Excel and filter by `category`, `change_type`, `confidence`, `needs_review` to triage what your scraper needs to update.
- **`rollup.md`** and **`rollup.csv`** — index of all diffs sorted by impact score so you can see which versions had the most disruptive changes.

The Scraping Impact section at the top of each diff report is the key output — it surfaces only critical changes (block moves, removes, reshapes, sheet renames) and warning changes (label edits, internal modifications), grouped by severity. This is what a human reviewer should look at first when deciding what scraping mappings to update.

## The stability classifier

The system learns which cells are **labels** (stable across versions) vs **values** (volatile across versions) by observing changes over time. Cells classified as VALUE have their changes silently suppressed because they're noise — version numbers, dates, totals that change every month by design. Cells classified as LABEL have their changes reported normally because they're identity anchors that scraping depends on.

**Bootstrap behavior** — for the first 2–3 diffs in a fresh database, the classifier has no history and reports everything with `needs_review=True`. By the 4th–5th diff, cells with enough observations get classified and noise drops dramatically. By month 8–10, nearly everything is classified.

**Escape hatch** — if a cell historically classified VALUE suddenly receives content that is short, alphabetic, and never seen in its history, the change is reported anyway with `needs_review=True`. This catches the rare case where a label gets renamed on a cell whose previous values looked like data.

**Behavioral guarantee** — when the ledger is empty or thin, the classifier returns UNKNOWN for everything and the differ falls back to reporting all changes with review flags. The classifier can only ever *remove* noise, never add it. Disabling the classifier entirely (set `USE_STABILITY_CLASSIFIER = False` in `config.py`) reverts to pre-classifier behavior with no other side effects.

## Tuning

All knobs live in `config.py`. The ones you'll most likely touch:

- **`GAP_TOLERANCE`** (default `0`) — how many empty rows/cols can sit between cells before block detection treats them as separate blocks. Lower = more, smaller blocks; higher = fewer, larger blocks. Run `batch --block-debug` to write per-file ASCII overlays of detected blocks; visually verify clustering quality and adjust.

- **`MAX_BLOCK_ROW_SPAN`** / **`MAX_BLOCK_COL_SPAN`** (default `50_000` / `600`) — sanity caps on block size, set generously to accommodate real data tables. Only pathological over-merging should trigger warnings.

- **`SKIP_BLOCK_DETECTION_SHEETS`** (default empty set) — sheets to bypass block detection entirely. Cells, merges, and named ranges are still extracted; only block-level tracking is skipped. Use for opaque data tables where block diffs add no value but produce noise every month. Example:
  ```python
  SKIP_BLOCK_DETECTION_SHEETS = {"DataTable", "RawExports", "Transactions"}
  ```

- **`TRACK_NAMED_RANGE_ADDRESS_CHANGES`** (default `False`) — when False, only named range additions and removals are reported. Address shifts on existing ranges are silently ignored to suppress noise from internal data table growth. Set to True if you need to audit address drift.

- **`AUTO_ACCEPT_THRESHOLD`** (default `0.85`) and **`REVIEW_ZONE_THRESHOLD`** (default `0.50`) — block matching score thresholds. Matches above auto-accept are taken silently; matches in the review zone are taken but flagged `needs_review=True`. Tune after seeing real data.

- **`STABILITY_LABEL_THRESHOLD`** (default `0.70`) and **`STABILITY_VALUE_THRESHOLD`** (default `0.30`) — classifier thresholds. A cell's stability score is `1 - (changes / transitions)`. Cells above `LABEL_THRESHOLD` become LABEL; below `VALUE_THRESHOLD` become VALUE; in between stay UNKNOWN.

- **`STABILITY_MIN_OBSERVATIONS`** (default `3`) — a cell needs this many observations before the classifier will emit LABEL or VALUE. Below this it stays UNKNOWN.

- **`USE_STABILITY_CLASSIFIER`** (default `True`) — master kill switch. Set to False to disable the classifier entirely.

The full set of impact weights, fuzzy thresholds, and matching weights is documented inline in `config.py`.

## Architecture

```
xlsb file
   │
   ▼  converter.py        (win32com → cached .xlsx, macros disabled)
xlsx file
   │
   ▼  extractor.py        (openpyxl → sparse cell map + named ranges + merges)
WorkbookSnapshot (sheets, cells, named ranges, merges) — no blocks yet
   │
   ▼  block_detector.py   (connected components + gap tolerance + fingerprints)
WorkbookSnapshot (now with blocks)
   │           │
   ▼           ▼  storage.py        (JSON file + SQLite row)
   │           snapshots/<file>.json
   │
   ▼  differ.py           (sheet diff → matcher.py block matching → block diff
   │                       → named ranges → merges → impact → classifier filter)
WorkbookDiff
   │           │
   ▼           ▼  storage.py        (JSON + SQLite)
   │           diffs/<old>__to__<new>.json
   │
   ▼  reporter.py         (Markdown + CSV)
   │           reports/<old>__to__<new>.md
   │           reports/<old>__to__<new>.csv
   │
   ▼  stability.py        (update ledger from cell-by-cell observations)
   │
   ▼  reporter.write_rollup_reports()
               reports/rollup.md
               reports/rollup.csv
```

## Project layout

```
20260410_ExcelEvolutionTracker/        ← project root
├── pyproject.toml
├── uv.lock
├── README.md
├── .venv/                             ← uv-managed
├── cache/xlsx/                        ← converted XLSX files (runtime)
├── snapshots/                         ← JSON snapshots (runtime)
├── diffs/                             ← JSON diffs (runtime)
├── reports/                           ← Markdown + CSV reports (runtime)
├── db/
│   └── evolution.db                   ← SQLite (runtime)
└── excel_evo_tracker/                 ← the package source code
    ├── __init__.py
    ├── cli.py
    ├── config.py
    ├── converter.py
    ├── extractor.py
    ├── block_detector.py
    ├── matcher.py
    ├── differ.py
    ├── stability.py
    ├── storage.py
    ├── reporter.py
    ├── pipeline.py
    ├── models.py
    └── db/
        └── schema.sql                 ← schema template (ships with package)
```

## Module reference

- **`config.py`** — paths, tunable thresholds, weights, kill switches
- **`models.py`** — Pydantic data models (Cell, Block, WorkbookSnapshot, WorkbookDiff, etc.)
- **`converter.py`** — XLSB → XLSX with hash-based caching and macro disabling
- **`extractor.py`** — XLSX → snapshot using openpyxl
- **`block_detector.py`** — block clustering with gap-tolerance merging, fingerprinting, debug overlay writer
- **`matcher.py`** — multi-pass block matching (fingerprint → composite → label-set fallback)
- **`differ.py`** — full diff engine with classifier integration and impact scoring
- **`stability.py`** — temporal cell classification ledger
- **`storage.py`** — JSON + SQLite persistence and queries
- **`reporter.py`** — Markdown + CSV writers, scraping-impact section, rollup, timeline
- **`pipeline.py`** — `run_batch`, `run_incremental`, `run_compare` orchestrators
- **`cli.py`** — argparse entry point with all subcommands
- **`db/schema.sql`** — SQLite schema (loaded by `init_db()`)
