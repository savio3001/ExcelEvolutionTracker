"""
Stability ledger — temporal cell classification.

Learns which cells are "labels" (stable across months) versus "values"
(volatile across months) by observing changes over time. A cell keyed
by (sheet_name, primary_label, rel_row, rel_col) accumulates
observation and transition counts in SQLite; after enough observations,
the classifier can confidently suppress value changes and report only
label changes.

Design guarantees:
  - Empty ledger → behaves identically to no-classifier mode (all
    changes reported with needs_review=True).
  - Classifier output is strictly LABEL | VALUE | UNKNOWN.
  - Option (b) escape hatch: if a cell classified VALUE receives a new
    value that is short, alphabetic-ish, and never seen in its history,
    report the change anyway with needs_review=True. Catches the
    "historically volatile cell suddenly gets renamed" edge case.
"""

from __future__ import annotations

import logging
from enum import Enum
from typing import Optional

from . import config
from .models import Block, WorkbookDiff, WorkbookSnapshot
from .storage import get_connection

logger = logging.getLogger(__name__)


class CellRole(str, Enum):
    LABEL = "label"
    VALUE = "value"
    UNKNOWN = "unknown"


# ── Schema management ─────────────────────────────────────────────────


STABILITY_SCHEMA = """
CREATE TABLE IF NOT EXISTS cell_stability (
    sheet_name        TEXT NOT NULL,
    primary_label     TEXT NOT NULL,
    rel_row           INTEGER NOT NULL,
    rel_col           INTEGER NOT NULL,
    observations      INTEGER NOT NULL DEFAULT 0,
    transitions       INTEGER NOT NULL DEFAULT 0,
    changes           INTEGER NOT NULL DEFAULT 0,
    last_value        TEXT,
    first_seen        TEXT,
    last_seen         TEXT,
    PRIMARY KEY (sheet_name, primary_label, rel_row, rel_col)
);

CREATE INDEX IF NOT EXISTS idx_stability_sheet ON cell_stability(sheet_name);

CREATE TABLE IF NOT EXISTS cell_value_history (
    sheet_name        TEXT NOT NULL,
    primary_label     TEXT NOT NULL,
    rel_row           INTEGER NOT NULL,
    rel_col           INTEGER NOT NULL,
    value             TEXT NOT NULL,
    seen_count        INTEGER NOT NULL DEFAULT 1,
    PRIMARY KEY (sheet_name, primary_label, rel_row, rel_col, value)
);
"""


def ensure_stability_schema() -> None:
    """Create stability tables if they don't exist. Idempotent."""
    with get_connection() as conn:
        conn.executescript(STABILITY_SCHEMA)


# ── Classification ────────────────────────────────────────────────────


def _stability_score(observations: int, transitions: int, changes: int) -> float:
    """Fraction of transitions where the cell did NOT change. 1.0 = stable."""
    if transitions == 0:
        return 1.0 if observations > 0 else 0.0
    return 1.0 - (changes / transitions)


def classify_cell(
    sheet_name: str,
    primary_label: str,
    rel_row: int,
    rel_col: int,
) -> tuple[CellRole, float, int]:
    """
    Look up a cell's learned classification.

    Returns (role, stability_score, observation_count). If the cell is
    below STABILITY_MIN_OBSERVATIONS or absent from the ledger entirely,
    returns UNKNOWN.
    """
    if not config.USE_STABILITY_CLASSIFIER:
        return CellRole.UNKNOWN, 0.0, 0

    ensure_stability_schema()
    with get_connection() as conn:
        row = conn.execute(
            """SELECT observations, transitions, changes FROM cell_stability
               WHERE sheet_name=? AND primary_label=? AND rel_row=? AND rel_col=?""",
            (sheet_name, primary_label, rel_row, rel_col),
        ).fetchone()

    if row is None or row["observations"] < config.STABILITY_MIN_OBSERVATIONS:
        return CellRole.UNKNOWN, 0.0, (row["observations"] if row else 0)

    score = _stability_score(row["observations"], row["transitions"], row["changes"])
    if score >= config.STABILITY_LABEL_THRESHOLD:
        return CellRole.LABEL, score, row["observations"]
    if score <= config.STABILITY_VALUE_THRESHOLD:
        return CellRole.VALUE, score, row["observations"]
    return CellRole.UNKNOWN, score, row["observations"]


def is_likely_label_rename(
    sheet_name: str,
    primary_label: str,
    rel_row: int,
    rel_col: int,
    new_value: str,
) -> bool:
    """
    Option (b) escape hatch: a cell classified VALUE may still be
    reporting a real label rename if its new value:
      - is short (≤ STABILITY_ESCAPE_MAX_LENGTH chars)
      - contains at least one letter
      - has never been seen in this cell's history
    """
    if not new_value or len(new_value) > config.STABILITY_ESCAPE_MAX_LENGTH:
        return False
    if not any(c.isalpha() for c in new_value):
        return False

    ensure_stability_schema()
    with get_connection() as conn:
        row = conn.execute(
            """SELECT 1 FROM cell_value_history
               WHERE sheet_name=? AND primary_label=? AND rel_row=? AND rel_col=?
                 AND value=? LIMIT 1""",
            (sheet_name, primary_label, rel_row, rel_col, new_value),
        ).fetchone()
    return row is None


# ── Ledger updates ────────────────────────────────────────────────────


def _block_cell_map(block: Block) -> dict[tuple[int, int], str]:
    """Map (rel_row, rel_col) → value for every label cell in a block."""
    tr, tc = block.top_left
    out: dict[tuple[int, int], str] = {}
    for c in block.label_cells:
        rel = (c.row - tr, c.col - tc)
        out[rel] = str(c.value) if c.value is not None else ""
    return out


def update_ledger(
    old_snap: WorkbookSnapshot,
    new_snap: WorkbookSnapshot,
    diff: WorkbookDiff,
) -> dict:
    """
    Walk matched block pairs in both snapshots and update the ledger.

    This is called after save_diff() from the pipeline. Idempotent at
    the transaction level but not across re-runs — re-running the same
    diff will double-count. Call once per new diff.
    """
    from .matcher import match_blocks

    ensure_stability_schema()

    updated = 0
    inserted = 0
    changes_seen = 0
    snapshot_date = new_snap.snapshot_date

    with get_connection() as conn:
        # For each sheet present in both snapshots, re-match blocks so
        # we know which pairs to trace. (The differ already did this, but
        # we don't carry the MatchPairs through the WorkbookDiff model.)
        common_sheets = set(old_snap.sheet_names) & set(new_snap.sheet_names)
        for sheet_name in common_sheets:
            old_sheet = old_snap.sheets[sheet_name]
            new_sheet = new_snap.sheets[sheet_name]
            if not old_sheet.blocks or not new_sheet.blocks:
                continue

            result = match_blocks(old_sheet.blocks, new_sheet.blocks)

            for pair in result.pairs:
                old_map = _block_cell_map(pair.old_block)
                new_map = _block_cell_map(pair.new_block)
                primary = pair.new_block.primary_label or pair.old_block.primary_label or "(unnamed)"

                all_positions = set(old_map) | set(new_map)
                for rel in all_positions:
                    old_val = old_map.get(rel)
                    new_val = new_map.get(rel)
                    if new_val is None:
                        # Cell disappeared — don't record, can't observe further
                        continue

                    changed = (old_val is not None and old_val != new_val)
                    had_prior = old_val is not None

                    # Upsert the stability row
                    conn.execute(
                        """INSERT INTO cell_stability
                             (sheet_name, primary_label, rel_row, rel_col,
                              observations, transitions, changes,
                              last_value, first_seen, last_seen)
                           VALUES (?, ?, ?, ?, 1, ?, ?, ?, ?, ?)
                           ON CONFLICT(sheet_name, primary_label, rel_row, rel_col) DO UPDATE SET
                             observations = observations + 1,
                             transitions = transitions + ?,
                             changes = changes + ?,
                             last_value = excluded.last_value,
                             last_seen = excluded.last_seen""",
                        (
                            sheet_name, primary, rel[0], rel[1],
                            1 if had_prior else 0,
                            1 if changed else 0,
                            new_val, snapshot_date, snapshot_date,
                            1 if had_prior else 0,
                            1 if changed else 0,
                        ),
                    )
                    if changed:
                        changes_seen += 1

                    # Record value in history (capped)
                    conn.execute(
                        """INSERT INTO cell_value_history
                             (sheet_name, primary_label, rel_row, rel_col, value, seen_count)
                           VALUES (?, ?, ?, ?, ?, 1)
                           ON CONFLICT(sheet_name, primary_label, rel_row, rel_col, value) DO UPDATE SET
                             seen_count = seen_count + 1""",
                        (sheet_name, primary, rel[0], rel[1], new_val),
                    )
                    updated += 1

    logger.info(
        "Ledger updated: %d cell observations recorded, %d changes seen",
        updated, changes_seen,
    )
    return {"observations": updated, "changes": changes_seen}


# ── Audit dump for CLI ────────────────────────────────────────────────


def dump_sheet_classification(sheet_name: str) -> list[dict]:
    """Return every tracked cell for a sheet with its current classification."""
    ensure_stability_schema()
    with get_connection() as conn:
        rows = conn.execute(
            """SELECT sheet_name, primary_label, rel_row, rel_col,
                      observations, transitions, changes, last_value,
                      first_seen, last_seen
               FROM cell_stability
               WHERE sheet_name = ?
               ORDER BY primary_label, rel_row, rel_col""",
            (sheet_name,),
        ).fetchall()

    out = []
    for r in rows:
        score = _stability_score(r["observations"], r["transitions"], r["changes"])
        if r["observations"] < config.STABILITY_MIN_OBSERVATIONS:
            role = CellRole.UNKNOWN
        elif score >= config.STABILITY_LABEL_THRESHOLD:
            role = CellRole.LABEL
        elif score <= config.STABILITY_VALUE_THRESHOLD:
            role = CellRole.VALUE
        else:
            role = CellRole.UNKNOWN
        out.append({
            **dict(r),
            "stability_score": round(score, 3),
            "role": role.value,
        })
    return out


def list_known_sheets() -> list[str]:
    """All sheets with at least one ledger entry."""
    ensure_stability_schema()
    with get_connection() as conn:
        rows = conn.execute(
            "SELECT DISTINCT sheet_name FROM cell_stability ORDER BY sheet_name"
        ).fetchall()
    return [r["sheet_name"] for r in rows]
