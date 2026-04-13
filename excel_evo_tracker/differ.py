"""
Diff engine — turns two WorkbookSnapshots into a WorkbookDiff.

Pipeline:
  1. Sheet-level diff (added / removed / renamed / reordered)
  2. Per matched-sheet pair → block matching → block-level diff
     (added / removed / moved / reshaped / label_changed / modified)
  3. Named range diff (added / removed / modified by exact name)
  4. Merged region diff per matched sheet (set diff)
  5. Summary counts + impact score

Typical usage:

    from excel_evo_tracker.differ import diff_snapshots
    diff = diff_snapshots(old_snap, new_snap)
"""

from __future__ import annotations

import datetime as dt
import logging

from . import config
from .matcher import MatchResult, match_blocks, match_sheet_names
from .models import (
    Block,
    BlockChange,
    ChangeType,
    DiffSummary,
    MergedRegionChange,
    NamedRange,
    NamedRangeChange,
    SheetChange,
    SheetSnapshot,
    WorkbookDiff,
    WorkbookSnapshot,
)

logger = logging.getLogger(__name__)


# ── Sheet-level diff ──────────────────────────────────────────────────


def _build_fingerprint_sets(snap: WorkbookSnapshot) -> dict[str, set[str]]:
    """Map sheet name → set of its block fingerprints."""
    return {
        name: {b.fingerprint for b in sheet.blocks}
        for name, sheet in snap.sheets.items()
    }


def _diff_sheets(
    old_snap: WorkbookSnapshot,
    new_snap: WorkbookSnapshot,
) -> tuple[list[SheetChange], dict[str, str]]:
    """
    Detect added, removed, renamed, and reordered sheets.

    Sheets listed in config.IGNORE_SHEETS are excluded from the diff
    even if they appear in one snapshot and not the other — they're
    simply treated as if they don't exist.

    Returns the list of SheetChanges plus the matches dict
    {old_name: new_name} for downstream block diffing.
    """
    changes: list[SheetChange] = []
    ignored = getattr(config, "IGNORE_SHEETS", set())

    old_names_filtered = [n for n in old_snap.sheet_names if n not in ignored]
    new_names_filtered = [n for n in new_snap.sheet_names if n not in ignored]

    old_fps = _build_fingerprint_sets(old_snap)
    new_fps = _build_fingerprint_sets(new_snap)
    # Drop ignored names from the fingerprint sets too
    old_fps = {k: v for k, v in old_fps.items() if k not in ignored}
    new_fps = {k: v for k, v in new_fps.items() if k not in ignored}

    matches, unmatched_old, unmatched_new = match_sheet_names(
        old_names_filtered,
        new_names_filtered,
        old_fps,
        new_fps,
    )

    # Renames (matched but with different names)
    for old_name, new_name in matches.items():
        if old_name != new_name:
            changes.append(SheetChange(
                change_type=ChangeType.RENAMED,
                sheet_name=new_name,
                old_name=old_name,
                new_name=new_name,
                detail=f"Sheet renamed from {old_name!r} to {new_name!r}",
            ))

    # Removed sheets
    for name in unmatched_old:
        old_idx = old_snap.sheet_names.index(name)
        changes.append(SheetChange(
            change_type=ChangeType.REMOVED,
            sheet_name=name,
            old_name=name,
            old_index=old_idx,
            detail=f"Sheet {name!r} was removed",
        ))

    # Added sheets
    for name in unmatched_new:
        new_idx = new_snap.sheet_names.index(name)
        changes.append(SheetChange(
            change_type=ChangeType.ADDED,
            sheet_name=name,
            new_name=name,
            new_index=new_idx,
            detail=f"Sheet {name!r} was added",
        ))

    # Reordering — only meaningful for sheets present in both
    for old_name, new_name in matches.items():
        old_idx = old_snap.sheet_names.index(old_name)
        new_idx = new_snap.sheet_names.index(new_name)
        if old_idx != new_idx:
            changes.append(SheetChange(
                change_type=ChangeType.REORDERED,
                sheet_name=new_name,
                old_name=old_name,
                new_name=new_name,
                old_index=old_idx,
                new_index=new_idx,
                detail=f"Sheet position changed: {old_idx} → {new_idx}",
            ))

    return changes, matches


# ── Block-level diff ──────────────────────────────────────────────────


def _diff_matched_block_pair(
    sheet_name: str,
    old: Block,
    new: Block,
    confidence: float,
    needs_review: bool,
) -> list[BlockChange]:
    """
    Compare a matched (old, new) block pair and emit BlockChanges.

    A single pair can produce multiple changes (e.g., a block that
    moved AND had its label edited AND gained an internal label).
    """
    changes: list[BlockChange] = []

    moved = old.top_left != new.top_left
    reshaped = old.shape != new.shape
    label_changed = (old.primary_label or "") != (new.primary_label or "")

    old_label_set = set(old.labels)
    new_label_set = set(new.labels)
    raw_added = sorted(new_label_set - old_label_set)
    raw_removed = sorted(old_label_set - new_label_set)

    # Consult stability classifier to filter volatile-value noise.
    # When the classifier has insufficient history, everything falls
    # through to UNKNOWN and gets reported with needs_review=True —
    # strictly no worse than pre-classifier behavior.
    added_labels, removed_labels, classifier_forced_review = _filter_by_stability(
        sheet_name=sheet_name,
        old_block=old,
        new_block=new,
        raw_added=raw_added,
        raw_removed=raw_removed,
    )
    internal_changed = bool(added_labels or removed_labels)

    if moved:
        delta = (
            new.top_left[0] - old.top_left[0],
            new.top_left[1] - old.top_left[1],
        )
        changes.append(BlockChange(
            change_type=ChangeType.MOVED,
            sheet_name=sheet_name,
            old_block=old,
            new_block=new,
            match_confidence=confidence,
            needs_review=needs_review,
            position_delta=delta,
            detail=(
                f"Block {old.primary_label!r} moved from "
                f"{_addr(old.top_left)} to {_addr(new.top_left)} "
                f"(Δrow={delta[0]:+d}, Δcol={delta[1]:+d})"
            ),
        ))

    if label_changed:
        changes.append(BlockChange(
            change_type=ChangeType.LABEL_CHANGED,
            sheet_name=sheet_name,
            old_block=old,
            new_block=new,
            match_confidence=confidence,
            needs_review=needs_review,
            label_diff={"old": old.primary_label, "new": new.primary_label},
            detail=(
                f"Primary label changed: {old.primary_label!r} → "
                f"{new.primary_label!r}"
            ),
        ))

    if reshaped:
        changes.append(BlockChange(
            change_type=ChangeType.RESHAPED,
            sheet_name=sheet_name,
            old_block=old,
            new_block=new,
            match_confidence=confidence,
            needs_review=needs_review,
            shape_diff={"old": list(old.shape), "new": list(new.shape)},
            detail=(
                f"Shape changed: {old.shape} → {new.shape} "
                f"(Δrows={new.shape[0] - old.shape[0]:+d}, "
                f"Δcols={new.shape[1] - old.shape[1]:+d})"
            ),
        ))

    if internal_changed:
        changes.append(BlockChange(
            change_type=ChangeType.MODIFIED,
            sheet_name=sheet_name,
            old_block=old,
            new_block=new,
            match_confidence=confidence,
            needs_review=needs_review or classifier_forced_review,
            internal_labels_added=added_labels,
            internal_labels_removed=removed_labels,
            detail=(
                f"Internal labels changed: "
                f"+{added_labels} -{removed_labels}"
            ),
        ))

    return changes


def _addr(pos: tuple[int, int]) -> str:
    """Convert (row, col) → 'B5' style address."""
    row, col = pos
    letters = ""
    c = col
    while c > 0:
        c, rem = divmod(c - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return f"{letters}{row}"


def _filter_by_stability(
    sheet_name: str,
    old_block,
    new_block,
    raw_added: list[str],
    raw_removed: list[str],
) -> tuple[list[str], list[str], bool]:
    """
    Filter raw added/removed label lists through the stability classifier.

    For each changed label, locates its cell in the block, asks the
    classifier about the cell's role, and either keeps the change,
    drops it (VALUE classification), or keeps it with needs_review=True
    (UNKNOWN classification, or VALUE classification that triggered the
    rename escape hatch).

    Returns (filtered_added, filtered_removed, any_forced_review).
    """
    if not config.USE_STABILITY_CLASSIFIER:
        return raw_added, raw_removed, False

    # Lazy import avoids circular dependency
    from .stability import CellRole, classify_cell, is_likely_label_rename

    primary = new_block.primary_label or old_block.primary_label or "(unnamed)"

    # Build (rel_row, rel_col) → value maps for both blocks' label cells
    def _cell_map(block):
        tr, tc = block.top_left
        return {
            (c.row - tr, c.col - tc): str(c.value) if c.value is not None else ""
            for c in block.label_cells
        }

    old_map = _cell_map(old_block)
    new_map = _cell_map(new_block)

    # Invert: value → position (first occurrence wins; duplicate labels in
    # a block are rare enough that we tolerate the approximation)
    old_value_to_pos = {}
    for pos, val in old_map.items():
        old_value_to_pos.setdefault(val, pos)
    new_value_to_pos = {}
    for pos, val in new_map.items():
        new_value_to_pos.setdefault(val, pos)

    filtered_added: list[str] = []
    filtered_removed: list[str] = []
    forced_review = False

    for label in raw_added:
        pos = new_value_to_pos.get(label)
        if pos is None:
            filtered_added.append(label)
            continue
        role, _score, _obs = classify_cell(sheet_name, primary, pos[0], pos[1])
        if role == CellRole.LABEL:
            filtered_added.append(label)
        elif role == CellRole.VALUE:
            if is_likely_label_rename(sheet_name, primary, pos[0], pos[1], label):
                filtered_added.append(label)
                forced_review = True
            # else: suppress silently
        else:  # UNKNOWN
            filtered_added.append(label)
            forced_review = True

    for label in raw_removed:
        pos = old_value_to_pos.get(label)
        if pos is None:
            filtered_removed.append(label)
            continue
        role, _score, _obs = classify_cell(sheet_name, primary, pos[0], pos[1])
        if role == CellRole.LABEL:
            filtered_removed.append(label)
        elif role == CellRole.VALUE:
            pass  # suppress silently
        else:  # UNKNOWN
            filtered_removed.append(label)
            forced_review = True

    return filtered_added, filtered_removed, forced_review


def _diff_blocks_for_sheet(
    sheet_name: str,
    old_sheet: SheetSnapshot,
    new_sheet: SheetSnapshot,
) -> list[BlockChange]:
    """Run block matcher on a sheet pair and convert results into changes."""
    result: MatchResult = match_blocks(old_sheet.blocks, new_sheet.blocks)
    changes: list[BlockChange] = []

    # Matched pairs → moves/reshapes/label changes/modifications
    for pair in result.pairs:
        pair_changes = _diff_matched_block_pair(
            sheet_name=sheet_name,
            old=pair.old_block,
            new=pair.new_block,
            confidence=pair.confidence,
            needs_review=pair.needs_review,
        )
        changes.extend(pair_changes)

    # Unmatched old → removed
    for old in result.unmatched_old:
        changes.append(BlockChange(
            change_type=ChangeType.REMOVED,
            sheet_name=sheet_name,
            old_block=old,
            match_confidence=0.0,
            detail=(
                f"Block {old.primary_label!r} at {_addr(old.top_left)} "
                f"was removed"
            ),
        ))

    # Unmatched new → added
    for new in result.unmatched_new:
        changes.append(BlockChange(
            change_type=ChangeType.ADDED,
            sheet_name=sheet_name,
            new_block=new,
            match_confidence=0.0,
            detail=(
                f"Block {new.primary_label!r} at {_addr(new.top_left)} "
                f"was added"
            ),
        ))

    return changes


# ── Named range diff ──────────────────────────────────────────────────


def _diff_named_ranges(
    old_ranges: list[NamedRange],
    new_ranges: list[NamedRange],
) -> list[NamedRangeChange]:
    """
    Set-diff named ranges by exact name.

    If config.TRACK_NAMED_RANGE_ADDRESS_CHANGES is False, refers_to and
    scope changes on existing named ranges are NOT reported — only
    additions and removals. This is the default because internal data
    table growth tends to shift range addresses every month, producing
    noisy diffs that don't affect scraping mappings.
    """
    changes: list[NamedRangeChange] = []

    old_by_name = {nr.name: nr for nr in old_ranges}
    new_by_name = {nr.name: nr for nr in new_ranges}

    old_names = set(old_by_name)
    new_names = set(new_by_name)

    for name in sorted(new_names - old_names):
        nr = new_by_name[name]
        changes.append(NamedRangeChange(
            change_type=ChangeType.ADDED,
            name=name,
            new_refers_to=nr.refers_to,
            new_scope=nr.scope,
            detail=f"Named range {name!r} added → {nr.refers_to}",
        ))

    for name in sorted(old_names - new_names):
        nr = old_by_name[name]
        changes.append(NamedRangeChange(
            change_type=ChangeType.REMOVED,
            name=name,
            old_refers_to=nr.refers_to,
            old_scope=nr.scope,
            detail=f"Named range {name!r} removed (was {nr.refers_to})",
        ))

    if config.TRACK_NAMED_RANGE_ADDRESS_CHANGES:
        for name in sorted(old_names & new_names):
            old_nr = old_by_name[name]
            new_nr = new_by_name[name]
            if old_nr.refers_to != new_nr.refers_to or old_nr.scope != new_nr.scope:
                changes.append(NamedRangeChange(
                    change_type=ChangeType.MODIFIED,
                    name=name,
                    old_refers_to=old_nr.refers_to,
                    new_refers_to=new_nr.refers_to,
                    old_scope=old_nr.scope,
                    new_scope=new_nr.scope,
                    detail=(
                        f"Named range {name!r}: "
                        f"{old_nr.refers_to} → {new_nr.refers_to}"
                    ),
                ))

    return changes


# ── Merged region diff ────────────────────────────────────────────────


def _diff_merged_regions(
    sheet_name: str,
    old_sheet: SheetSnapshot,
    new_sheet: SheetSnapshot,
) -> list[MergedRegionChange]:
    """Set-diff merged ranges per sheet."""
    changes: list[MergedRegionChange] = []
    old_set = set(old_sheet.merged_regions)
    new_set = set(new_sheet.merged_regions)

    for region in sorted(new_set - old_set):
        changes.append(MergedRegionChange(
            change_type=ChangeType.ADDED,
            sheet_name=sheet_name,
            region=region,
        ))
    for region in sorted(old_set - new_set):
        changes.append(MergedRegionChange(
            change_type=ChangeType.REMOVED,
            sheet_name=sheet_name,
            region=region,
        ))
    return changes


# ── Summary & impact scoring ──────────────────────────────────────────


def _build_summary(diff: WorkbookDiff) -> DiffSummary:
    """Aggregate change counts into the DiffSummary model."""
    s = DiffSummary()

    for sc in diff.sheet_changes:
        if sc.change_type == ChangeType.ADDED:
            s.sheets_added += 1
        elif sc.change_type == ChangeType.REMOVED:
            s.sheets_removed += 1
        elif sc.change_type == ChangeType.RENAMED:
            s.sheets_renamed += 1
        elif sc.change_type == ChangeType.REORDERED:
            s.sheets_reordered += 1

    for bc in diff.block_changes:
        if bc.needs_review:
            s.needs_review_count += 1
        if bc.change_type == ChangeType.ADDED:
            s.blocks_added += 1
        elif bc.change_type == ChangeType.REMOVED:
            s.blocks_removed += 1
        elif bc.change_type == ChangeType.MOVED:
            s.blocks_moved += 1
        elif bc.change_type == ChangeType.RESHAPED:
            s.blocks_reshaped += 1
        elif bc.change_type == ChangeType.LABEL_CHANGED:
            s.blocks_label_changed += 1
        elif bc.change_type == ChangeType.MODIFIED:
            s.blocks_modified += 1

    for nrc in diff.named_range_changes:
        if nrc.change_type == ChangeType.ADDED:
            s.named_ranges_added += 1
        elif nrc.change_type == ChangeType.REMOVED:
            s.named_ranges_removed += 1
        elif nrc.change_type == ChangeType.MODIFIED:
            s.named_ranges_modified += 1

    for mc in diff.merged_region_changes:
        if mc.change_type == ChangeType.ADDED:
            s.merges_added += 1
        elif mc.change_type == ChangeType.REMOVED:
            s.merges_removed += 1

    s.total_changes = (
        len(diff.sheet_changes)
        + len(diff.block_changes)
        + len(diff.named_range_changes)
        + len(diff.merged_region_changes)
    )
    return s


def _compute_impact_score(summary: DiffSummary) -> float:
    """Weighted sum of change counts, normalized to [0, 1]."""
    raw = (
        summary.sheets_added * config.IMPACT_WEIGHTS["sheet_added"]
        + summary.sheets_removed * config.IMPACT_WEIGHTS["sheet_removed"]
        + summary.sheets_renamed * config.IMPACT_WEIGHTS["sheet_renamed"]
        + summary.blocks_added * config.IMPACT_WEIGHTS["block_added"]
        + summary.blocks_removed * config.IMPACT_WEIGHTS["block_removed"]
        + summary.blocks_moved * config.IMPACT_WEIGHTS["block_moved"]
        + summary.blocks_reshaped * config.IMPACT_WEIGHTS["block_reshaped"]
        + summary.blocks_label_changed * config.IMPACT_WEIGHTS["block_label_changed"]
        + summary.blocks_modified * config.IMPACT_WEIGHTS["block_modified"]
        + summary.named_ranges_added * config.IMPACT_WEIGHTS["named_range_added"]
        + summary.named_ranges_removed * config.IMPACT_WEIGHTS["named_range_removed"]
        + summary.named_ranges_modified * config.IMPACT_WEIGHTS["named_range_modified"]
        + summary.merges_added * config.IMPACT_WEIGHTS["merge_added"]
        + summary.merges_removed * config.IMPACT_WEIGHTS["merge_removed"]
    )
    return min(1.0, raw / config.IMPACT_NORMALIZATION)


# ── Public API ────────────────────────────────────────────────────────


def diff_snapshots(
    old_snap: WorkbookSnapshot,
    new_snap: WorkbookSnapshot,
) -> WorkbookDiff:
    """
    Compute a structural diff between two workbook snapshots.

    Both snapshots must have already had block detection run on them
    (i.e., each sheet's `blocks` field is populated). The differ does
    not call the block detector itself.
    """
    diff = WorkbookDiff(
        old_file=old_snap.file_name,
        new_file=new_snap.file_name,
        old_month=old_snap.month_label,
        new_month=new_snap.month_label,
        diff_date=dt.datetime.now(dt.timezone.utc).isoformat(),
    )

    # 1. Sheet-level diff
    sheet_changes, sheet_matches = _diff_sheets(old_snap, new_snap)
    diff.sheet_changes = sheet_changes

    # 2. Block-level + merged region diff per matched sheet pair
    for old_name, new_name in sheet_matches.items():
        old_sheet = old_snap.sheets[old_name]
        new_sheet = new_snap.sheets[new_name]

        block_changes = _diff_blocks_for_sheet(new_name, old_sheet, new_sheet)
        diff.block_changes.extend(block_changes)

        merge_changes = _diff_merged_regions(new_name, old_sheet, new_sheet)
        diff.merged_region_changes.extend(merge_changes)

    # 3. Named range diff (workbook level)
    diff.named_range_changes = _diff_named_ranges(
        old_snap.named_ranges, new_snap.named_ranges
    )

    # 4. Summary + impact score
    diff.summary = _build_summary(diff)
    diff.impact_score = _compute_impact_score(diff.summary)

    logger.info(
        "Diff %s → %s: %d total changes, impact %.2f",
        old_snap.file_name, new_snap.file_name,
        diff.summary.total_changes, diff.impact_score,
    )
    return diff
