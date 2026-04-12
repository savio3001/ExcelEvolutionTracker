"""
Central configuration for the Excel Evolution Tracker.

All tunable parameters live here so they can be adjusted without touching
business logic. Paths are relative to the project root by default but can
be overridden via environment variables or CLI args.
"""

from pathlib import Path
import os

# ── Paths ─────────────────────────────────────────────────────────────

PROJECT_ROOT = Path(os.environ.get("EVO_ROOT", Path(__file__).parent.parent)).resolve()

# Input: where raw XLSB files live (user-provided or CLI override)
XLSB_INPUT_DIR = PROJECT_ROOT / "xlsb_input"

# Cache: converted XLSX files (one-time conversion, reused thereafter)
XLSX_CACHE_DIR = PROJECT_ROOT / "cache" / "xlsx"
CONVERSION_MANIFEST = XLSX_CACHE_DIR / "_manifest.json"

# Output: snapshots, diffs, reports
SNAPSHOT_DIR = PROJECT_ROOT / "snapshots"
DIFF_DIR = PROJECT_ROOT / "diffs"
REPORT_DIR = PROJECT_ROOT / "reports"

# SQLite database
DB_PATH = PROJECT_ROOT / "db" / "evolution.db"
DB_SCHEMA_PATH = Path(__file__).parent / "db" / "schema.sql"

# Ensure runtime directories exist
for _d in (XLSX_CACHE_DIR, SNAPSHOT_DIR, DIFF_DIR, REPORT_DIR, DB_PATH.parent):
    _d.mkdir(parents=True, exist_ok=True)


# ── Block detection tuning ────────────────────────────────────────────

# Max number of empty rows/cols between cells that should still be
# considered part of the same block. 0 = tight clustering (any gap
# splits). Recommended for dense templates with data tables.
GAP_TOLERANCE = 1

# Ignore clusters smaller than this (noise suppression)
MIN_BLOCK_CELLS = 1

# A block must contain at least this many label (text) cells to be tracked
MIN_LABEL_CELLS = 1

# Maximum span (rows or cols) of a single block — guards against
# accidentally merging unrelated regions. Set generously so legitimate
# large data tables don't trigger warnings; only pathological over-
# merging should exceed these.
MAX_BLOCK_ROW_SPAN = 50_000
MAX_BLOCK_COL_SPAN = 600

# Sheets to skip block detection entirely — they'll still be extracted
# (cells, merges, named ranges all captured) but no blocks will be
# emitted. Use for opaque data tables where block-level tracking adds
# no value. Match by exact sheet name; case-sensitive.
SKIP_BLOCK_DETECTION_SHEETS: set[str] = set()

# How far around a block to look when collecting neighbor_context
NEIGHBOR_RADIUS_ROWS = 10
NEIGHBOR_RADIUS_COLS = 5


# ── Matching tuning ───────────────────────────────────────────────────

# Weights used to compute composite match score (must sum to 1.0)
MATCH_WEIGHTS = {
    "label_similarity": 0.40,
    "fingerprint_match": 0.25,
    "position_proximity": 0.20,
    "neighbor_context": 0.10,
    "shape_similarity": 0.05,
}
assert abs(sum(MATCH_WEIGHTS.values()) - 1.0) < 1e-9, "MATCH_WEIGHTS must sum to 1.0"

# Score thresholds for match acceptance
AUTO_ACCEPT_THRESHOLD = 0.85   # score >= this → auto-match
REVIEW_ZONE_THRESHOLD = 0.50   # score in [this, AUTO_ACCEPT) → flagged for review
# score < REVIEW_ZONE_THRESHOLD → rejected, treated as add/remove

# Fuzzy string matching threshold (rapidfuzz ratio, 0-100) for pass 2
FUZZY_LABEL_THRESHOLD = 70

# Sheet rename detection thresholds
SHEET_NAME_FUZZY_THRESHOLD = 60       # min fuzzy name similarity to consider rename
SHEET_CONTENT_OVERLAP_THRESHOLD = 0.5  # min fingerprint set overlap to confirm rename


# ── Impact scoring weights ────────────────────────────────────────────

# How much each change type contributes to the impact score. Higher =
# more likely to break scraping mappings.
IMPACT_WEIGHTS = {
    "sheet_added": 0.30,
    "sheet_removed": 0.30,
    "sheet_renamed": 0.20,
    "block_added": 0.15,
    "block_removed": 0.20,
    "block_moved": 0.10,
    "block_reshaped": 0.10,
    "block_label_changed": 0.10,
    "block_modified": 0.05,
    "named_range_added": 0.10,
    "named_range_removed": 0.10,
    "named_range_modified": 0.10,
    "merge_added": 0.05,
    "merge_removed": 0.05,
}

# Normalization divisor — raw score is divided by this then clamped to [0, 1].
# Tune after seeing real data; a typical month might score 0.5-2.0 pre-norm.
IMPACT_NORMALIZATION = 2.0


# ── Stability classifier (learned label/value classification) ────────

# Master kill switch — when False, classifier is bypassed entirely and
# behavior reverts to Phase 4 (all changes reported, none suppressed).
USE_STABILITY_CLASSIFIER = True

# A cell must have at least this many monthly observations before the
# classifier will emit LABEL or VALUE (below this → UNKNOWN).
STABILITY_MIN_OBSERVATIONS = 3

# Stability score thresholds. A cell's score = 1 - (changes / transitions).
# ≥ LABEL_THRESHOLD → LABEL (never changes, safe to track)
# ≤ VALUE_THRESHOLD → VALUE (always changes, noise — suppress)
# in between    → UNKNOWN (flag for review)
STABILITY_LABEL_THRESHOLD = 0.70
STABILITY_VALUE_THRESHOLD = 0.30

# Option-b escape hatch: a VALUE-classified cell whose new content
# looks like a rename (short, alphabetic, never seen before) gets
# reported anyway with needs_review=True.
STABILITY_ESCAPE_MAX_LENGTH = 30


# ── Named range tracking ──────────────────────────────────────────────

# If True, track changes to a named range's refers_to (address). Set
# this False if internal data table growth causes noisy address-shift
# diffs every month — when False, only additions and removals are
# reported, which is usually enough for scraping mapping maintenance.
TRACK_NAMED_RANGE_ADDRESS_CHANGES = False


# ── Extraction behavior ───────────────────────────────────────────────

# Include numeric cells in the snapshot? They're not used for label
# matching but help block detection recognize table shapes.
INCLUDE_NUMERIC_CELLS = True

# Include hidden sheets?
INCLUDE_HIDDEN_SHEETS = True

# Max cells per sheet before we warn (very large sheets may be slow)
LARGE_SHEET_WARN_THRESHOLD = 100_000


# ── Report generation ─────────────────────────────────────────────────

# Report output formats
GENERATE_MARKDOWN_REPORT = True
GENERATE_CSV_REPORT = True

# When reports get very large (many changes), truncate detail sections
# and link to the full JSON diff instead
REPORT_MAX_CHANGES_PER_SECTION = 500
