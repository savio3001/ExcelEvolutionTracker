"""
Persistence layer for snapshots and diffs.

Two storage backends working together:
  - JSON files on disk hold the full WorkbookSnapshot and WorkbookDiff
    payloads (lossless, human-readable, git-friendly)
  - SQLite holds indexed metadata for fast cross-month queries

The JSON files are the source of truth; the database is a queryable
index that points to them. If the DB is ever lost it can be rebuilt
by re-importing the JSON files.

Typical usage:

    from excel_evo_tracker.storage import (
        save_snapshot, save_diff, get_latest_snapshot,
        list_diffs_by_impact, init_db,
    )
    init_db()
    save_snapshot(snap)
    save_diff(diff)
"""

from __future__ import annotations

import json
import logging
import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Optional

from . import config
from .models import (
    BlockChange,
    ChangeType,
    MergedRegionChange,
    NamedRangeChange,
    SheetChange,
    WorkbookDiff,
    WorkbookSnapshot,
)

logger = logging.getLogger(__name__)


# ── Database connection management ────────────────────────────────────


@contextmanager
def get_connection() -> Iterator[sqlite3.Connection]:
    """
    Yield a SQLite connection with foreign keys enabled and row_factory
    set so query results behave like dicts.
    """
    config.DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(config.DB_PATH))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db(force: bool = False) -> None:
    """
    Create database tables from schema.sql.

    Idempotent — re-running on an existing database is safe and won't
    drop data. Pass force=True to delete the existing database first.
    """
    if force and config.DB_PATH.exists():
        logger.warning("Deleting existing database at %s", config.DB_PATH)
        config.DB_PATH.unlink()

    schema_path = config.DB_SCHEMA_PATH
    if not schema_path.exists():
        # Schema file missing — write a minimal inline version as a fallback
        raise FileNotFoundError(
            f"Schema file not found: {schema_path}. "
            f"Make sure db/schema.sql is present."
        )

    schema_sql = schema_path.read_text()
    with get_connection() as conn:
        conn.executescript(schema_sql)
    logger.info("Database initialized at %s", config.DB_PATH)


# ── JSON paths ────────────────────────────────────────────────────────


def _snapshot_json_path(snap: WorkbookSnapshot) -> Path:
    """Pick a stable on-disk filename for a snapshot."""
    label = snap.month_label or Path(snap.file_name).stem
    safe = "".join(c if c.isalnum() or c in "-_." else "_" for c in label)
    return config.SNAPSHOT_DIR / f"{safe}.json"


def _diff_json_path(diff: WorkbookDiff) -> Path:
    """Pick a stable on-disk filename for a diff."""
    old_label = diff.old_month or Path(diff.old_file).stem
    new_label = diff.new_month or Path(diff.new_file).stem
    safe_old = "".join(c if c.isalnum() or c in "-_." else "_" for c in old_label)
    safe_new = "".join(c if c.isalnum() or c in "-_." else "_" for c in new_label)
    return config.DIFF_DIR / f"{safe_old}__to__{safe_new}.json"


# ── Snapshot save / load ──────────────────────────────────────────────


def save_snapshot(snap: WorkbookSnapshot) -> Path:
    """
    Persist a snapshot to disk and index it in the database.

    Returns the path to the JSON file. If a snapshot with the same
    file_hash already exists in the DB, the JSON is overwritten and
    the DB row is updated rather than duplicated.
    """
    json_path = _snapshot_json_path(snap)
    json_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.write_text(
        snap.model_dump_json(indent=config.JSON_INDENT),
        encoding="utf-8",
    )

    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO snapshots (
                file_name, month_label, file_hash, snapshot_date,
                sheet_count, block_count, cell_count, named_range_count,
                json_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(file_hash) DO UPDATE SET
                file_name = excluded.file_name,
                month_label = excluded.month_label,
                snapshot_date = excluded.snapshot_date,
                sheet_count = excluded.sheet_count,
                block_count = excluded.block_count,
                cell_count = excluded.cell_count,
                named_range_count = excluded.named_range_count,
                json_path = excluded.json_path
            """,
            (
                snap.file_name,
                snap.month_label,
                snap.file_hash,
                snap.snapshot_date,
                len(snap.sheet_names),
                snap.total_blocks,
                snap.total_cells,
                len(snap.named_ranges),
                str(json_path),
            ),
        )

    logger.info("Saved snapshot: %s (hash %s)", json_path.name, snap.file_hash[:8])
    return json_path


def load_snapshot(json_path: Path) -> WorkbookSnapshot:
    """Load a snapshot from its JSON file on disk."""
    json_path = Path(json_path)
    return WorkbookSnapshot.model_validate_json(json_path.read_text(encoding="utf-8"))


def get_snapshot_by_hash(file_hash: str) -> Optional[WorkbookSnapshot]:
    """Look up a snapshot by file hash, or None if not present."""
    with get_connection() as conn:
        row = conn.execute(
            "SELECT json_path FROM snapshots WHERE file_hash = ?",
            (file_hash,),
        ).fetchone()
    if row is None:
        return None
    return load_snapshot(Path(row["json_path"]))


def get_latest_snapshot() -> Optional[WorkbookSnapshot]:
    """Return the most recent snapshot by snapshot_date (or None)."""
    with get_connection() as conn:
        row = conn.execute(
            "SELECT json_path FROM snapshots ORDER BY snapshot_date DESC LIMIT 1"
        ).fetchone()
    if row is None:
        return None
    return load_snapshot(Path(row["json_path"]))


def list_snapshots() -> list[dict]:
    """List all snapshots in the database, ordered by snapshot_date."""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, file_name, month_label, snapshot_date,
                   sheet_count, block_count, cell_count, named_range_count, json_path
            FROM snapshots
            ORDER BY snapshot_date ASC
            """
        ).fetchall()
    return [dict(r) for r in rows]


# ── Diff save / load ──────────────────────────────────────────────────


def _change_to_db_row(
    diff_summary_id: int,
    category: str,
    change_type: str,
    sheet_name: Optional[str],
    element_name: Optional[str],
    old_value: Optional[str],
    new_value: Optional[str],
    confidence: float,
    needs_review: bool,
    impact_weight: float,
    detail: str,
) -> tuple:
    return (
        diff_summary_id,
        category,
        change_type,
        sheet_name,
        element_name,
        old_value,
        new_value,
        confidence,
        1 if needs_review else 0,
        impact_weight,
        detail,
    )


def _flatten_changes_to_rows(
    diff: WorkbookDiff, diff_summary_id: int
) -> list[tuple]:
    """Flatten every change in a diff into row tuples ready for executemany."""
    from . import config as cfg

    rows: list[tuple] = []

    for sc in diff.sheet_changes:
        weight = cfg.IMPACT_WEIGHTS.get(f"sheet_{sc.change_type.value}", 0.0)
        rows.append(_change_to_db_row(
            diff_summary_id, "sheet", sc.change_type.value,
            sc.sheet_name, sc.sheet_name,
            json.dumps({"old_name": sc.old_name, "old_index": sc.old_index}),
            json.dumps({"new_name": sc.new_name, "new_index": sc.new_index}),
            confidence=1.0, needs_review=False,
            impact_weight=weight, detail=sc.detail,
        ))

    for bc in diff.block_changes:
        weight = cfg.IMPACT_WEIGHTS.get(f"block_{bc.change_type.value}", 0.0)
        old_summary = None
        new_summary = None
        if bc.old_block:
            old_summary = json.dumps({
                "primary_label": bc.old_block.primary_label,
                "top_left": list(bc.old_block.top_left),
                "shape": list(bc.old_block.shape),
                "labels": bc.old_block.labels,
            })
        if bc.new_block:
            new_summary = json.dumps({
                "primary_label": bc.new_block.primary_label,
                "top_left": list(bc.new_block.top_left),
                "shape": list(bc.new_block.shape),
                "labels": bc.new_block.labels,
            })
        element_name = (
            (bc.new_block and bc.new_block.primary_label)
            or (bc.old_block and bc.old_block.primary_label)
            or "(unnamed)"
        )
        rows.append(_change_to_db_row(
            diff_summary_id, "block", bc.change_type.value,
            bc.sheet_name, element_name,
            old_summary, new_summary,
            confidence=bc.match_confidence,
            needs_review=bc.needs_review,
            impact_weight=weight, detail=bc.detail,
        ))

    for nrc in diff.named_range_changes:
        weight = cfg.IMPACT_WEIGHTS.get(f"named_range_{nrc.change_type.value}", 0.0)
        rows.append(_change_to_db_row(
            diff_summary_id, "named_range", nrc.change_type.value,
            None, nrc.name,
            json.dumps({"refers_to": nrc.old_refers_to, "scope": nrc.old_scope}),
            json.dumps({"refers_to": nrc.new_refers_to, "scope": nrc.new_scope}),
            confidence=1.0, needs_review=False,
            impact_weight=weight, detail=nrc.detail,
        ))

    for mc in diff.merged_region_changes:
        weight = cfg.IMPACT_WEIGHTS.get(f"merge_{mc.change_type.value}", 0.0)
        rows.append(_change_to_db_row(
            diff_summary_id, "merge", mc.change_type.value,
            mc.sheet_name, mc.region,
            None, None,
            confidence=1.0, needs_review=False,
            impact_weight=weight,
            detail=f"Merged region {mc.region} {mc.change_type.value} on {mc.sheet_name}",
        ))

    return rows


def save_diff(
    diff: WorkbookDiff,
    md_report_path: Optional[Path] = None,
    csv_report_path: Optional[Path] = None,
) -> Path:
    """
    Persist a diff to disk and index it in the database.

    Both snapshots referenced by the diff must already exist in the DB
    (their hashes are looked up to populate the foreign keys).

    Args:
        diff: The WorkbookDiff to persist.
        md_report_path: Optional path to a Markdown report (recorded in DB).
        csv_report_path: Optional path to a CSV report (recorded in DB).

    Returns:
        Path to the JSON file.
    """
    json_path = _diff_json_path(diff)
    json_path.parent.mkdir(parents=True, exist_ok=True)
    json_path.write_text(
        diff.model_dump_json(indent=config.JSON_INDENT),
        encoding="utf-8",
    )

    with get_connection() as conn:
        # Look up snapshot IDs by file_name + month_label
        old_row = conn.execute(
            "SELECT id FROM snapshots WHERE file_name = ? AND "
            "(month_label = ? OR (month_label IS NULL AND ? IS NULL)) "
            "ORDER BY id DESC LIMIT 1",
            (diff.old_file, diff.old_month, diff.old_month),
        ).fetchone()
        new_row = conn.execute(
            "SELECT id FROM snapshots WHERE file_name = ? AND "
            "(month_label = ? OR (month_label IS NULL AND ? IS NULL)) "
            "ORDER BY id DESC LIMIT 1",
            (diff.new_file, diff.new_month, diff.new_month),
        ).fetchone()

        if old_row is None or new_row is None:
            raise RuntimeError(
                f"Cannot save diff: snapshots not in DB for "
                f"{diff.old_file!r} or {diff.new_file!r}. "
                f"Save snapshots before saving diffs."
            )

        old_id = old_row["id"]
        new_id = new_row["id"]

        # Upsert diff_summaries
        cursor = conn.execute(
            """
            INSERT INTO diff_summaries (
                old_snapshot_id, new_snapshot_id,
                old_file, new_file, old_month, new_month,
                diff_date, impact_score, total_changes,
                needs_review_count, summary_json, json_path,
                md_report_path, csv_report_path
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(old_snapshot_id, new_snapshot_id) DO UPDATE SET
                diff_date = excluded.diff_date,
                impact_score = excluded.impact_score,
                total_changes = excluded.total_changes,
                needs_review_count = excluded.needs_review_count,
                summary_json = excluded.summary_json,
                json_path = excluded.json_path,
                md_report_path = excluded.md_report_path,
                csv_report_path = excluded.csv_report_path
            """,
            (
                old_id, new_id,
                diff.old_file, diff.new_file,
                diff.old_month, diff.new_month,
                diff.diff_date,
                diff.impact_score,
                diff.summary.total_changes,
                diff.summary.needs_review_count,
                diff.summary.model_dump_json(),
                str(json_path),
                str(md_report_path) if md_report_path else None,
                str(csv_report_path) if csv_report_path else None,
            ),
        )

        # Get the diff_summary_id for the upserted row
        ds_row = conn.execute(
            "SELECT id FROM diff_summaries WHERE old_snapshot_id = ? AND new_snapshot_id = ?",
            (old_id, new_id),
        ).fetchone()
        diff_summary_id = ds_row["id"]

        # Replace any existing change rows for this diff
        conn.execute("DELETE FROM changes WHERE diff_summary_id = ?", (diff_summary_id,))
        change_rows = _flatten_changes_to_rows(diff, diff_summary_id)
        if change_rows:
            conn.executemany(
                """
                INSERT INTO changes (
                    diff_summary_id, change_category, change_type,
                    sheet_name, element_name, old_value, new_value,
                    confidence, needs_review, impact_weight, detail
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                change_rows,
            )

    logger.info(
        "Saved diff: %s (%d changes, impact %.2f)",
        json_path.name, diff.summary.total_changes, diff.impact_score,
    )
    return json_path


def load_diff(json_path: Path) -> WorkbookDiff:
    """Load a diff from its JSON file on disk."""
    json_path = Path(json_path)
    return WorkbookDiff.model_validate_json(json_path.read_text(encoding="utf-8"))


def list_diffs() -> list[dict]:
    """List all diffs in the database, ordered by diff_date."""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, old_file, new_file, old_month, new_month, diff_date,
                   impact_score, total_changes, needs_review_count,
                   json_path, md_report_path, csv_report_path
            FROM diff_summaries
            ORDER BY diff_date ASC
            """
        ).fetchall()
    return [dict(r) for r in rows]


def list_diffs_by_impact(limit: int = 10) -> list[dict]:
    """Return the top-N diffs ranked by impact score, descending."""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, old_file, new_file, old_month, new_month,
                   impact_score, total_changes, needs_review_count, json_path
            FROM diff_summaries
            ORDER BY impact_score DESC
            LIMIT ?
            """,
            (limit,),
        ).fetchall()
    return [dict(r) for r in rows]


def find_changes_for_element(element_name: str) -> list[dict]:
    """
    Find every recorded change involving an element of a given name.

    Useful for tracing the history of a single block, sheet, or named
    range across all months. Matches by exact element_name.
    """
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT c.*, ds.old_month, ds.new_month, ds.diff_date,
                   ds.old_file, ds.new_file
            FROM changes c
            JOIN diff_summaries ds ON c.diff_summary_id = ds.id
            WHERE c.element_name = ?
            ORDER BY ds.diff_date ASC
            """,
            (element_name,),
        ).fetchall()
    return [dict(r) for r in rows]


def find_changes_needing_review() -> list[dict]:
    """Return every change flagged needs_review across all diffs."""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT c.*, ds.old_month, ds.new_month, ds.diff_date,
                   ds.old_file, ds.new_file
            FROM changes c
            JOIN diff_summaries ds ON c.diff_summary_id = ds.id
            WHERE c.needs_review = 1
            ORDER BY ds.diff_date ASC
            """
        ).fetchall()
    return [dict(r) for r in rows]


def find_block_timeline(
    sheet_name: str,
    primary_label: str,
    *,
    limit: int | None = None,
) -> list[dict]:
    """
    Return every recorded block-level change involving a specific
    (sheet, primary_label) pair, oldest first.

    Args:
        sheet_name: Exact sheet name to filter on.
        primary_label: Exact block primary label to filter on.
        limit: If provided, return only the most recent N changes
            (still ordered oldest-first in the result list).
    """
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT c.*, ds.old_month, ds.new_month, ds.diff_date,
                   ds.old_file, ds.new_file
            FROM changes c
            JOIN diff_summaries ds ON c.diff_summary_id = ds.id
            WHERE c.change_category = 'block'
              AND c.sheet_name = ?
              AND c.element_name = ?
            ORDER BY ds.diff_date ASC
            """,
            (sheet_name, primary_label),
        ).fetchall()
    result = [dict(r) for r in rows]
    if limit is not None and limit > 0 and len(result) > limit:
        result = result[-limit:]
    return result


def list_known_block_labels(sheet_name: str) -> list[str]:
    """All distinct block primary labels recorded for a sheet."""
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT DISTINCT element_name FROM changes
            WHERE change_category = 'block' AND sheet_name = ?
            ORDER BY element_name
            """,
            (sheet_name,),
        ).fetchall()
    return [r["element_name"] for r in rows if r["element_name"]]
