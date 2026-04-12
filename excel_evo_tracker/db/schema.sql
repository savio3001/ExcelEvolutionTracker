-- Excel Evolution Tracker — SQLite schema
--
-- Three tables:
--   snapshots:      one row per workbook snapshot
--   changes:        one row per individual change (block, sheet, named range, merge)
--   diff_summaries: one row per workbook-pair diff with aggregated counts
--
-- The full snapshot/diff payloads live as JSON files on disk; this DB
-- stores indexed metadata for fast cross-month querying.

CREATE TABLE IF NOT EXISTS snapshots (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name           TEXT    NOT NULL,
    month_label         TEXT,
    file_hash           TEXT    NOT NULL,
    snapshot_date       TEXT    NOT NULL,
    sheet_count         INTEGER NOT NULL,
    block_count         INTEGER NOT NULL,
    cell_count          INTEGER NOT NULL,
    named_range_count   INTEGER NOT NULL,
    json_path           TEXT    NOT NULL,
    UNIQUE(file_hash)
);

CREATE INDEX IF NOT EXISTS idx_snapshots_month ON snapshots(month_label);
CREATE INDEX IF NOT EXISTS idx_snapshots_date  ON snapshots(snapshot_date);

CREATE TABLE IF NOT EXISTS diff_summaries (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    old_snapshot_id     INTEGER NOT NULL REFERENCES snapshots(id),
    new_snapshot_id     INTEGER NOT NULL REFERENCES snapshots(id),
    old_file            TEXT    NOT NULL,
    new_file            TEXT    NOT NULL,
    old_month           TEXT,
    new_month           TEXT,
    diff_date           TEXT    NOT NULL,
    impact_score        REAL    NOT NULL,
    total_changes       INTEGER NOT NULL,
    needs_review_count  INTEGER NOT NULL,
    summary_json        TEXT    NOT NULL,
    json_path           TEXT    NOT NULL,
    md_report_path      TEXT,
    csv_report_path     TEXT,
    UNIQUE(old_snapshot_id, new_snapshot_id)
);

CREATE INDEX IF NOT EXISTS idx_diff_summaries_impact ON diff_summaries(impact_score);
CREATE INDEX IF NOT EXISTS idx_diff_summaries_date   ON diff_summaries(diff_date);

CREATE TABLE IF NOT EXISTS changes (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    diff_summary_id     INTEGER NOT NULL REFERENCES diff_summaries(id),
    change_category     TEXT    NOT NULL,   -- sheet | block | named_range | merge
    change_type         TEXT    NOT NULL,   -- added | removed | moved | renamed | reshaped | label_changed | modified | reordered
    sheet_name          TEXT,
    element_name        TEXT,               -- block primary_label, sheet name, or named range name
    old_value           TEXT,               -- JSON snippet of relevant old state
    new_value           TEXT,               -- JSON snippet of relevant new state
    confidence          REAL,
    needs_review        INTEGER NOT NULL DEFAULT 0,
    impact_weight       REAL,
    detail              TEXT
);

CREATE INDEX IF NOT EXISTS idx_changes_diff       ON changes(diff_summary_id);
CREATE INDEX IF NOT EXISTS idx_changes_element    ON changes(element_name);
CREATE INDEX IF NOT EXISTS idx_changes_type       ON changes(change_category, change_type);
CREATE INDEX IF NOT EXISTS idx_changes_review     ON changes(needs_review);
