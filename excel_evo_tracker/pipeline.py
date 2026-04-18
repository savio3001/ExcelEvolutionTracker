"""
Pipeline orchestrator.

Two entry points:

  run_batch(xlsb_dir):
      Convert every XLSB in the directory, extract snapshots, detect
      blocks, generate sequential diffs (file_1 → file_2 → file_3 …),
      persist everything, and write reports. Used for the initial
      one-time backfill over your existing 70 files.

  run_incremental(new_xlsb):
      Convert one new file, extract its snapshot, diff it against the
      most recent stored snapshot, and append the result. Used in
      production for the monthly cadence.

The "process one file" path is shared between the two modes via the
internal _process_file() helper.
"""

from __future__ import annotations

import logging
import re
from pathlib import Path
from typing import Optional

from . import config
from .block_detector import detect_blocks_for_workbook, write_debug_report
from .converter import convert_single, convert_batch
from .differ import diff_snapshots
from .extractor import extract_workbook
from .models import WorkbookDiff, WorkbookSnapshot
from .reporter import write_diff_reports, write_rollup_reports
from .storage import (
    get_latest_snapshot,
    init_db,
    list_snapshots,
    load_snapshot,
    save_diff,
    save_snapshot,
)

logger = logging.getLogger(__name__)


# ── Filename → month label heuristics ─────────────────────────────────


_MONTH_PATTERNS = [
    re.compile(r"(\d{4})[-_]?(\d{2})"),         # 2024-01, 2024_01, 202401
    re.compile(r"(\d{2})[-_](\d{4})"),          # 01-2024
]


def _infer_month_label(path: Path) -> Optional[str]:
    """
    Try to extract a YYYY-MM month label from a filename.

    Returns None if no clear pattern is found — caller can fall back
    to the file stem or manual labeling.
    """
    stem = path.stem
    for pat in _MONTH_PATTERNS:
        m = pat.search(stem)
        if m:
            g1, g2 = m.group(1), m.group(2)
            # Decide which group is year vs month
            if len(g1) == 4:
                year, month = g1, g2
            else:
                year, month = g2, g1
            try:
                month_int = int(month)
                if 1 <= month_int <= 12:
                    return f"{year}-{int(month):02d}"
            except ValueError:
                continue
    return None


# ── Single-file processing ────────────────────────────────────────────


def _process_file(
    xlsb_path: Path,
    *,
    month_label: Optional[str] = None,
    write_block_debug: bool = False,
) -> WorkbookSnapshot:
    """
    Convert → extract → detect blocks → save. Shared by batch and incremental.
    """
    xlsx_path = convert_single(xlsb_path)

    if month_label is None:
        month_label = _infer_month_label(xlsb_path)

    snap = extract_workbook(
        xlsx_path,
        original_xlsb_path=xlsb_path,
        month_label=month_label,
    )
    detect_blocks_for_workbook(snap)
    save_snapshot(snap)

    if write_block_debug:
        write_debug_report(snap)

    return snap


# ── Batch mode ────────────────────────────────────────────────────────


def run_batch(
    xlsb_dir: Path | None = None,
    *,
    pattern: str = "*.xlsb",
    write_block_debug: bool = False,
    write_rollup: bool = True,
    reset_db: bool = False,
) -> dict:
    """
    Process every XLSB in a directory and generate sequential diffs.

    Streaming design: at any moment only two snapshots are held in memory
    (the previous and the current). After each diff, the previous
    snapshot is released and the garbage collector is invoked before
    loading the next file. This keeps peak memory bounded regardless of
    batch size — 70 files cost the same RAM as 2.

    Args:
        xlsb_dir: Source directory. Defaults to config.XLSB_INPUT_DIR.
        pattern: Glob pattern for source files.
        write_block_debug: Also write per-file block detection debug
            reports (useful when tuning GAP_TOLERANCE).
        write_rollup: Write the cross-diff rollup index at the end.
        reset_db: Drop and recreate the database before processing.

    Returns:
        Summary dict with counts and lists of generated artifacts.
    """
    import gc

    xlsb_dir = Path(xlsb_dir) if xlsb_dir else config.XLSB_INPUT_DIR
    if not xlsb_dir.is_dir():
        raise NotADirectoryError(f"Not a directory: {xlsb_dir}")

    init_db(force=reset_db)

    xlsb_files = sorted(xlsb_dir.glob(pattern))
    if not xlsb_files:
        logger.warning("No files matching %r in %s", pattern, xlsb_dir)
        return {"snapshots": 0, "diffs": 0, "errors": []}

    # Pre-sort by filename so diffs go in chronological order regardless
    # of the filesystem's returned order. (When month_label isn't in the
    # filename, the alphanumeric sort on file name is typically equivalent.)
    logger.info("Batch processing %d files from %s", len(xlsb_files), xlsb_dir)

    errors: list[str] = []
    diffs_written: list[Path] = []
    snap_count = 0
    prev_snap: WorkbookSnapshot | None = None

    if config.USE_STABILITY_CLASSIFIER:
        from .stability import update_ledger as _update_ledger
    else:
        _update_ledger = None

    for idx, xlsb in enumerate(xlsb_files, start=1):
        # ── Process the current file into a snapshot ─────────────────
        try:
            curr_snap = _process_file(xlsb, write_block_debug=write_block_debug)
            snap_count += 1
        except Exception as e:
            logger.exception("Failed to process %s", xlsb.name)
            errors.append(f"{xlsb.name}: {e}")
            continue

        logger.info("[%d/%d] Processed %s", idx, len(xlsb_files), xlsb.name)

        # ── Diff against previous, write reports, update ledger ──────
        if prev_snap is not None:
            try:
                diff = diff_snapshots(prev_snap, curr_snap)
                md_path, csv_path = write_diff_reports(diff)
                save_diff(diff, md_report_path=md_path, csv_report_path=csv_path)
                diffs_written.append(md_path)

                if _update_ledger is not None:
                    _update_ledger(prev_snap, curr_snap, diff)

                # Release the diff object before dropping prev_snap —
                # diffs hold references to Block objects in both snapshots.
                del diff
            except Exception as e:
                logger.exception("Failed to diff %s → %s",
                                 prev_snap.file_name, curr_snap.file_name)
                errors.append(f"diff {prev_snap.file_name}→{curr_snap.file_name}: {e}")

        # ── Release previous snapshot, advance the window ────────────
        if prev_snap is not None:
            del prev_snap
        prev_snap = curr_snap
        del curr_snap

        # Force reclamation of the released snapshot's cells/blocks
        # before loading the next (potentially huge) workbook.
        gc.collect()

    # Release the final snapshot too
    if prev_snap is not None:
        del prev_snap
        gc.collect()

    if write_rollup and diffs_written:
        write_rollup_reports()

    return {
        "snapshots": snap_count,
        "diffs": len(diffs_written),
        "errors": errors,
        "report_dir": str(config.REPORT_DIR),
    }


# ── Incremental mode ──────────────────────────────────────────────────


def run_incremental(
    new_xlsb: Path,
    *,
    month_label: Optional[str] = None,
    write_block_debug: bool = False,
    write_rollup: bool = True,
) -> dict:
    """
    Process a single new file and diff it against the most recent stored snapshot.

    Args:
        new_xlsb: Path to the new XLSB file (typically the latest month).
        month_label: Optional explicit label, e.g. "2024-12". If omitted,
            the pipeline tries to infer one from the filename.
        write_block_debug: Also write a block detection debug report.
        write_rollup: Refresh the cross-diff rollup index after saving.

    Returns:
        Summary dict with the diff path (or None if this was the first file).
    """
    new_xlsb = Path(new_xlsb)
    init_db(force=False)

    prev_snap = get_latest_snapshot()
    new_snap = _process_file(
        new_xlsb,
        month_label=month_label,
        write_block_debug=write_block_debug,
    )

    if prev_snap is None:
        logger.info("No previous snapshot — this is the first file in the database.")
        return {
            "snapshot": new_snap.file_name,
            "diff": None,
            "is_first": True,
        }

    diff = diff_snapshots(prev_snap, new_snap)
    md_path, csv_path = write_diff_reports(diff)
    save_diff(diff, md_report_path=md_path, csv_report_path=csv_path)

    if config.USE_STABILITY_CLASSIFIER:
        from .stability import update_ledger
        update_ledger(prev_snap, new_snap, diff)

    if write_rollup:
        write_rollup_reports()

    return {
        "snapshot": new_snap.file_name,
        "previous": prev_snap.file_name,
        "diff": str(md_path),
        "csv": str(csv_path),
        "impact_score": diff.impact_score,
        "total_changes": diff.summary.total_changes,
        "needs_review": diff.summary.needs_review_count,
    }


# ── Compare-only mode (no DB writes) ──────────────────────────────────


def run_compare(
    old_xlsb: Path,
    new_xlsb: Path,
    *,
    write_reports: bool = True,
) -> WorkbookDiff:
    """
    Compare two XLSB files directly and return the WorkbookDiff.

    This is a "side-by-side" mode that doesn't touch the database — useful
    for ad-hoc investigation or for testing changes to detection tuning.
    Reports are still written to REPORT_DIR if write_reports=True.
    """
    old_xlsx = convert_single(Path(old_xlsb))
    new_xlsx = convert_single(Path(new_xlsb))

    old_snap = extract_workbook(
        old_xlsx, original_xlsb_path=Path(old_xlsb),
        month_label=_infer_month_label(Path(old_xlsb)),
    )
    new_snap = extract_workbook(
        new_xlsx, original_xlsb_path=Path(new_xlsb),
        month_label=_infer_month_label(Path(new_xlsb)),
    )
    detect_blocks_for_workbook(old_snap)
    detect_blocks_for_workbook(new_snap)

    diff = diff_snapshots(old_snap, new_snap)
    if write_reports:
        write_diff_reports(diff)
    return diff


# ── Replace mode ──────────────────────────────────────────────────────


def run_replace(
    old_file_name: str,
    new_xlsb: Path,
    *,
    month_label: str | None = None,
    write_rollup: bool = True,
) -> dict:
    """
    Replace a snapshot in the database with a corrected version.

    1. Finds the old snapshot's neighbors (prev, next).
    2. Purges the old snapshot and all its associated diffs/reports.
    3. Processes the new file into a fresh snapshot.
    4. Regenerates diffs: prev→new and new→next (if neighbors exist).

    Args:
        old_file_name: The file_name of the snapshot to replace (as
            stored in the database, e.g. "UPT 1.02.23 - Dev.xlsb").
        new_xlsb: Path to the replacement XLSB file.
        month_label: Optional month label for the new snapshot.

    Returns:
        Summary dict describing what was purged and regenerated.
    """
    import gc
    from .storage import (
        get_snapshot_neighbors,
        load_snapshot,
        purge_snapshot,
        save_diff,
        save_snapshot,
    )

    new_xlsb = Path(new_xlsb)
    init_db(force=False)

    # Step 1: Find neighbors BEFORE purging
    prev_info, next_info = get_snapshot_neighbors(old_file_name)

    # Step 2: Purge the old snapshot and its diffs
    purged = purge_snapshot(old_file_name)
    if purged["snapshot_rows"] == 0:
        return {"error": f"No snapshot found for file {old_file_name!r}"}

    logger.info(
        "Purged snapshot %r: %d diffs, %d changes, %d files removed",
        old_file_name, purged["diff_rows"], purged["change_rows"], len(purged["files"]),
    )

    # Step 3: Process the new file
    new_snap = _process_file(new_xlsb, month_label=month_label)

    # Step 4: Regenerate diffs with neighbors
    diffs_generated = 0

    if config.USE_STABILITY_CLASSIFIER:
        from .stability import update_ledger as _update_ledger
    else:
        _update_ledger = None

    # prev → new
    if prev_info:
        try:
            prev_snap = load_snapshot(Path(prev_info["json_path"]))
            diff = diff_snapshots(prev_snap, new_snap)
            md_path, csv_path = write_diff_reports(diff)
            save_diff(diff, md_report_path=md_path, csv_report_path=csv_path)
            if _update_ledger:
                _update_ledger(prev_snap, new_snap, diff)
            diffs_generated += 1
            del diff, prev_snap
        except Exception as e:
            logger.exception("Failed to diff %s → %s", prev_info["file_name"], new_snap.file_name)

    # new → next
    if next_info:
        try:
            next_snap = load_snapshot(Path(next_info["json_path"]))
            diff = diff_snapshots(new_snap, next_snap)
            md_path, csv_path = write_diff_reports(diff)
            save_diff(diff, md_report_path=md_path, csv_report_path=csv_path)
            if _update_ledger:
                _update_ledger(new_snap, next_snap, diff)
            diffs_generated += 1
            del diff, next_snap
        except Exception as e:
            logger.exception("Failed to diff %s → %s", new_snap.file_name, next_info["file_name"])

    del new_snap
    gc.collect()

    if write_rollup:
        write_rollup_reports()

    return {
        "replaced": old_file_name,
        "new_file": new_xlsb.name,
        "purged": purged,
        "diffs_regenerated": diffs_generated,
        "prev_neighbor": prev_info["file_name"] if prev_info else None,
        "next_neighbor": next_info["file_name"] if next_info else None,
    }
