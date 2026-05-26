"""
Report writers — Markdown and CSV.

For each diff we produce two files in REPORT_DIR:

  <old>__to__<new>.md   — human-readable Markdown report with a
                          scraping-impact section at the top, then a
                          full breakdown by category. Sections are
                          truncated past REPORT_MAX_CHANGES_PER_SECTION
                          to keep huge diffs manageable; the full list
                          is always available in the JSON / CSV.

  <old>__to__<new>.csv  — one row per change. Columns are designed for
                          filtering and sorting in Excel: category,
                          change_type, sheet, element, old/new info,
                          confidence, needs_review, impact_weight, detail.

A rollup writer produces a single index.md / rollup.csv summarizing all
diffs in the project, sorted by impact score, so you can see at a glance
which months had the most disruptive changes.

Typical usage:

    from excel_evo_tracker.reporter import write_diff_reports, write_rollup_reports
    md_path, csv_path = write_diff_reports(diff)
    write_rollup_reports()
"""

from __future__ import annotations

import csv
import logging
from pathlib import Path
from typing import Optional

from . import config
from .models import (
    BlockChange,
    ChangeType,
    MergedRegionChange,
    NamedRangeChange,
    SheetChange,
    WorkbookDiff,
)
from .storage import list_diffs

logger = logging.getLogger(__name__)


# Change types that materially affect scraping mappings.
# These rise to the top of the markdown report so a human reviewer
# can decide what to update first.
CRITICAL_BLOCK_TYPES = {
    ChangeType.REMOVED,
    ChangeType.MOVED,
    ChangeType.RESHAPED,
}
WARNING_BLOCK_TYPES = {
    ChangeType.LABEL_CHANGED,
    ChangeType.MODIFIED,
    ChangeType.ADDED,
}
CRITICAL_SHEET_TYPES = {
    ChangeType.REMOVED,
    ChangeType.RENAMED,
}
CRITICAL_NAMED_RANGE_TYPES = {
    ChangeType.REMOVED,
    ChangeType.MODIFIED,
}


# ── Path helpers ──────────────────────────────────────────────────────


# ── Label display helper ──────────────────────────────────────────────


def _display_label(file_name: str | None, month_label: str | None) -> str:
    """
    Pick the best human-readable label for a snapshot or diff side.

    Prefers the file name (stem only — strip the .xlsb/.xlsx extension)
    because file names are unambiguous and require no inference. Falls
    back to month_label only if file_name is empty/missing.
    """
    if file_name:
        return Path(file_name).stem
    if month_label:
        return month_label
    return "?"


def _report_paths(diff: WorkbookDiff) -> tuple[Path, Path]:
    """Return (md_path, csv_path) for a given diff."""
    old_label = diff.old_month or Path(diff.old_file).stem
    new_label = diff.new_month or Path(diff.new_file).stem
    safe_old = "".join(c if c.isalnum() or c in "-_." else "_" for c in old_label)
    safe_new = "".join(c if c.isalnum() or c in "-_." else "_" for c in new_label)
    base = f"{safe_old}__to__{safe_new}"
    return (
        config.REPORT_DIR / f"{base}.md",
        config.REPORT_DIR / f"{base}.csv",
    )


def _addr(pos: tuple[int, int]) -> str:
    """Convert (row, col) → 'B5' style address."""
    row, col = pos
    letters = ""
    c = col
    while c > 0:
        c, rem = divmod(c - 1, 26)
        letters = chr(ord("A") + rem) + letters
    return f"{letters}{row}"


# ── Markdown rendering ────────────────────────────────────────────────


def _impact_bar(score: float, width: int = 20) -> str:
    """Render an ASCII progress bar from a 0-1 score."""
    filled = int(round(score * width))
    return "█" * filled + "░" * (width - filled)


def _render_scraping_impact_section(diff: WorkbookDiff) -> list[str]:
    """
    Top-of-report section: only the changes that materially affect
    scraping mappings, grouped by severity. This is what a human
    reviewer should look at first.
    """
    lines: list[str] = []

    critical: list[str] = []
    warnings: list[str] = []

    # Critical sheet changes
    for sc in diff.sheet_changes:
        if sc.change_type in CRITICAL_SHEET_TYPES:
            if sc.change_type == ChangeType.REMOVED:
                critical.append(f"- **Sheet removed**: `{sc.old_name}`")
            elif sc.change_type == ChangeType.RENAMED:
                critical.append(f"- **Sheet renamed**: `{sc.old_name}` → `{sc.new_name}`")

    # Critical block changes
    for bc in diff.block_changes:
        if bc.change_type in CRITICAL_BLOCK_TYPES:
            label = (
                (bc.old_block and bc.old_block.primary_label)
                or (bc.new_block and bc.new_block.primary_label)
                or "(unnamed)"
            )
            review_tag = " ⚠️ REVIEW" if bc.needs_review else ""
            if bc.change_type == ChangeType.MOVED:
                critical.append(
                    f"- **Block moved** [{bc.sheet_name}]: `{label}` "
                    f"{_addr(bc.old_block.top_left)} → {_addr(bc.new_block.top_left)}"
                    f" (Δ row={bc.position_delta[0]:+d}, col={bc.position_delta[1]:+d})"
                    f"{review_tag}"
                )
            elif bc.change_type == ChangeType.RESHAPED:
                critical.append(
                    f"- **Block reshaped** [{bc.sheet_name}]: `{label}` "
                    f"{tuple(bc.shape_diff['old'])} → {tuple(bc.shape_diff['new'])}"
                    f"{review_tag}"
                )
            elif bc.change_type == ChangeType.REMOVED:
                critical.append(
                    f"- **Block removed** [{bc.sheet_name}]: `{label}` "
                    f"at {_addr(bc.old_block.top_left)}"
                )
        elif bc.change_type in WARNING_BLOCK_TYPES:
            label = (
                (bc.new_block and bc.new_block.primary_label)
                or (bc.old_block and bc.old_block.primary_label)
                or "(unnamed)"
            )
            review_tag = " ⚠️ REVIEW" if bc.needs_review else ""
            if bc.change_type == ChangeType.LABEL_CHANGED:
                old_l = bc.label_diff.get("old", "?")
                new_l = bc.label_diff.get("new", "?")
                warnings.append(
                    f"- **Label changed** [{bc.sheet_name}]: "
                    f"`{old_l}` → `{new_l}`{review_tag}"
                )
            elif bc.change_type == ChangeType.MODIFIED:
                added = bc.internal_labels_added
                removed = bc.internal_labels_removed
                parts = []
                if added:
                    parts.append(f"added {added}")
                if removed:
                    parts.append(f"removed {removed}")
                warnings.append(
                    f"- **Internal labels changed** [{bc.sheet_name}] in `{label}`: "
                    f"{'; '.join(parts)}{review_tag}"
                )
            elif bc.change_type == ChangeType.ADDED:
                warnings.append(
                    f"- **New block** [{bc.sheet_name}]: `{label}` "
                    f"at {_addr(bc.new_block.top_left)}"
                )

    # Critical named range changes
    for nrc in diff.named_range_changes:
        if nrc.change_type in CRITICAL_NAMED_RANGE_TYPES:
            if nrc.change_type == ChangeType.REMOVED:
                critical.append(
                    f"- **Named range removed**: `{nrc.name}` "
                    f"(was `{nrc.old_refers_to}`)"
                )
            elif nrc.change_type == ChangeType.MODIFIED:
                critical.append(
                    f"- **Named range modified**: `{nrc.name}` "
                    f"`{nrc.old_refers_to}` → `{nrc.new_refers_to}`"
                )

    lines.append("## 🎯 Scraping Impact")
    lines.append("")
    if not critical and not warnings:
        lines.append("_No changes detected that affect scraping mappings._")
        lines.append("")
        return lines

    if critical:
        lines.append("### CRITICAL — likely breaks scraping")
        lines.append("")
        lines.extend(critical)
        lines.append("")
    if warnings:
        lines.append("### WARNING — may affect lookup or assumptions")
        lines.append("")
        lines.extend(warnings)
        lines.append("")
    return lines


def _truncate(items: list, limit: int) -> tuple[list, int]:
    """Return (visible_items, hidden_count)."""
    if len(items) <= limit:
        return items, 0
    return items[:limit], len(items) - limit


def _render_full_breakdown(diff: WorkbookDiff) -> list[str]:
    """All change categories in full detail (truncated past the limit)."""
    lines: list[str] = []
    limit = config.REPORT_MAX_CHANGES_PER_SECTION

    # Sheet changes
    lines.append("## Sheet Changes")
    lines.append("")
    if not diff.sheet_changes:
        lines.append("_No sheet-level changes._")
    else:
        visible, hidden = _truncate(diff.sheet_changes, limit)
        lines.append("| Type | Sheet | Detail |")
        lines.append("|------|-------|--------|")
        for sc in visible:
            lines.append(
                f"| {sc.change_type.value} | {sc.sheet_name} | {_md_escape(sc.detail)} |"
            )
        if hidden:
            lines.append(f"\n_…{hidden} more sheet changes truncated. See JSON / CSV for full list._")
    lines.append("")

    # Block changes — grouped by sheet for readability
    lines.append("## Block Changes")
    lines.append("")
    if not diff.block_changes:
        lines.append("_No block-level changes._")
    else:
        by_sheet: dict[str, list[BlockChange]] = {}
        for bc in diff.block_changes:
            by_sheet.setdefault(bc.sheet_name, []).append(bc)
        for sheet_name in sorted(by_sheet):
            sheet_changes = by_sheet[sheet_name]
            lines.append(f"### Sheet: `{sheet_name}` ({len(sheet_changes)} changes)")
            lines.append("")
            visible, hidden = _truncate(sheet_changes, limit)
            lines.append("| Type | Element | Confidence | Review | Detail |")
            lines.append("|------|---------|------------|--------|--------|")
            for bc in visible:
                label = (
                    (bc.new_block and bc.new_block.primary_label)
                    or (bc.old_block and bc.old_block.primary_label)
                    or "(unnamed)"
                )
                conf = f"{bc.match_confidence:.2f}" if bc.match_confidence else "—"
                review = "⚠️" if bc.needs_review else ""
                lines.append(
                    f"| {bc.change_type.value} | `{label}` | {conf} | {review} | "
                    f"{_md_escape(bc.detail)} |"
                )
            if hidden:
                lines.append(f"\n_…{hidden} more block changes in this sheet truncated._")
            lines.append("")

    # Named range changes
    lines.append("## Named Range Changes")
    lines.append("")
    if not diff.named_range_changes:
        lines.append("_No named range changes._")
    else:
        visible, hidden = _truncate(diff.named_range_changes, limit)
        lines.append("| Type | Name | Old refers_to | New refers_to |")
        lines.append("|------|------|---------------|---------------|")
        for nrc in visible:
            lines.append(
                f"| {nrc.change_type.value} | `{nrc.name}` | "
                f"`{nrc.old_refers_to or ''}` | `{nrc.new_refers_to or ''}` |"
            )
        if hidden:
            lines.append(f"\n_…{hidden} more named range changes truncated._")
    lines.append("")

    # Merged region changes
    lines.append("## Merged Region Changes")
    lines.append("")
    if not diff.merged_region_changes:
        lines.append("_No merged region changes._")
    else:
        by_sheet_m: dict[str, list[MergedRegionChange]] = {}
        for mc in diff.merged_region_changes:
            by_sheet_m.setdefault(mc.sheet_name, []).append(mc)
        for sheet_name in sorted(by_sheet_m):
            mc_list = by_sheet_m[sheet_name]
            lines.append(f"**{sheet_name}** ({len(mc_list)} changes)")
            for mc in mc_list[:limit]:
                marker = "+" if mc.change_type == ChangeType.ADDED else "−"
                lines.append(f"- {marker} `{mc.region}`")
            if len(mc_list) > limit:
                lines.append(f"- _…{len(mc_list) - limit} more truncated_")
            lines.append("")

    return lines


def _md_escape(text: str) -> str:
    """Escape pipes in table cells."""
    return text.replace("|", "\\|").replace("\n", " ")


def write_markdown_report(diff: WorkbookDiff, output_path: Optional[Path] = None) -> Path:
    """Write a Markdown report file for a diff. Returns the path."""
    if output_path is None:
        output_path, _ = _report_paths(diff)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    s = diff.summary
    lines: list[str] = []

    old_lbl = _display_label(diff.old_file, diff.old_month)
    new_lbl = _display_label(diff.new_file, diff.new_month)

    # Header
    lines.append(f"# Diff Report: {old_lbl} → {new_lbl}")
    lines.append("")
    lines.append(f"**Generated:** {diff.diff_date}")
    lines.append(f"**Files:** `{diff.old_file}` → `{diff.new_file}`")
    if diff.old_month or diff.new_month:
        lines.append(f"**Months:** `{diff.old_month or '?'}` → `{diff.new_month or '?'}`")
    lines.append("")
    lines.append(f"**Impact score:** `{diff.impact_score:.3f}`  `{_impact_bar(diff.impact_score)}`")
    lines.append(f"**Total changes:** {s.total_changes}")
    lines.append(f"**Needs review:** {s.needs_review_count}")
    lines.append("")

    # Summary table
    lines.append("## Summary")
    lines.append("")
    lines.append("| Category | Count |")
    lines.append("|----------|------:|")
    lines.append(f"| Sheets added | {s.sheets_added} |")
    lines.append(f"| Sheets removed | {s.sheets_removed} |")
    lines.append(f"| Sheets renamed | {s.sheets_renamed} |")
    lines.append(f"| Sheets reordered | {s.sheets_reordered} |")
    lines.append(f"| Blocks added | {s.blocks_added} |")
    lines.append(f"| Blocks removed | {s.blocks_removed} |")
    lines.append(f"| Blocks moved | {s.blocks_moved} |")
    lines.append(f"| Blocks reshaped | {s.blocks_reshaped} |")
    lines.append(f"| Blocks label-changed | {s.blocks_label_changed} |")
    lines.append(f"| Blocks modified (internal) | {s.blocks_modified} |")
    lines.append(f"| Named ranges added | {s.named_ranges_added} |")
    lines.append(f"| Named ranges removed | {s.named_ranges_removed} |")
    lines.append(f"| Named ranges modified | {s.named_ranges_modified} |")
    lines.append(f"| Merges added | {s.merges_added} |")
    lines.append(f"| Merges removed | {s.merges_removed} |")
    lines.append("")

    # Scraping impact (top-of-report attention grabber)
    lines.extend(_render_scraping_impact_section(diff))

    # Full breakdown
    lines.extend(_render_full_breakdown(diff))

    output_path.write_text("\n".join(lines), encoding="utf-8")
    logger.info("Wrote Markdown report: %s", output_path)
    return output_path


# ── CSV rendering ─────────────────────────────────────────────────────


CSV_COLUMNS = [
    "category",
    "change_type",
    "sheet",
    "element",
    "old_position",
    "new_position",
    "old_shape",
    "new_shape",
    "old_value",
    "new_value",
    "confidence",
    "needs_review",
    "impact_weight",
    "detail",
]


def _block_change_to_csv_row(bc: BlockChange) -> dict:
    label = (
        (bc.new_block and bc.new_block.primary_label)
        or (bc.old_block and bc.old_block.primary_label)
        or "(unnamed)"
    )
    weight = config.IMPACT_WEIGHTS.get(f"block_{bc.change_type.value}", 0.0)
    return {
        "category": "block",
        "change_type": bc.change_type.value,
        "sheet": bc.sheet_name,
        "element": label,
        "old_position": _addr(bc.old_block.top_left) if bc.old_block else "",
        "new_position": _addr(bc.new_block.top_left) if bc.new_block else "",
        "old_shape": str(bc.old_block.shape) if bc.old_block else "",
        "new_shape": str(bc.new_block.shape) if bc.new_block else "",
        "old_value": (bc.label_diff or {}).get("old", ""),
        "new_value": (bc.label_diff or {}).get("new", ""),
        "confidence": f"{bc.match_confidence:.3f}" if bc.match_confidence else "",
        "needs_review": "Y" if bc.needs_review else "",
        "impact_weight": f"{weight:.3f}",
        "detail": bc.detail,
    }


def _sheet_change_to_csv_row(sc: SheetChange) -> dict:
    weight = config.IMPACT_WEIGHTS.get(f"sheet_{sc.change_type.value}", 0.0)
    return {
        "category": "sheet",
        "change_type": sc.change_type.value,
        "sheet": sc.sheet_name,
        "element": sc.sheet_name,
        "old_position": "",
        "new_position": "",
        "old_shape": "",
        "new_shape": "",
        "old_value": sc.old_name or "",
        "new_value": sc.new_name or "",
        "confidence": "",
        "needs_review": "",
        "impact_weight": f"{weight:.3f}",
        "detail": sc.detail,
    }


def _named_range_change_to_csv_row(nrc: NamedRangeChange) -> dict:
    weight = config.IMPACT_WEIGHTS.get(f"named_range_{nrc.change_type.value}", 0.0)
    return {
        "category": "named_range",
        "change_type": nrc.change_type.value,
        "sheet": "",
        "element": nrc.name,
        "old_position": "",
        "new_position": "",
        "old_shape": "",
        "new_shape": "",
        "old_value": nrc.old_refers_to or "",
        "new_value": nrc.new_refers_to or "",
        "confidence": "",
        "needs_review": "",
        "impact_weight": f"{weight:.3f}",
        "detail": nrc.detail,
    }


def _merge_change_to_csv_row(mc: MergedRegionChange) -> dict:
    weight = config.IMPACT_WEIGHTS.get(f"merge_{mc.change_type.value}", 0.0)
    return {
        "category": "merge",
        "change_type": mc.change_type.value,
        "sheet": mc.sheet_name,
        "element": mc.region,
        "old_position": "",
        "new_position": "",
        "old_shape": "",
        "new_shape": "",
        "old_value": "",
        "new_value": "",
        "confidence": "",
        "needs_review": "",
        "impact_weight": f"{weight:.3f}",
        "detail": f"Merge {mc.change_type.value}: {mc.region}",
    }


def write_csv_report(diff: WorkbookDiff, output_path: Optional[Path] = None) -> Path:
    """Write a CSV with one row per change. Returns the path."""
    if output_path is None:
        _, output_path = _report_paths(diff)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    rows: list[dict] = []
    rows.extend(_sheet_change_to_csv_row(sc) for sc in diff.sheet_changes)
    rows.extend(_block_change_to_csv_row(bc) for bc in diff.block_changes)
    rows.extend(_named_range_change_to_csv_row(nrc) for nrc in diff.named_range_changes)
    rows.extend(_merge_change_to_csv_row(mc) for mc in diff.merged_region_changes)

    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)

    logger.info("Wrote CSV report: %s (%d rows)", output_path, len(rows))
    return output_path


# ── Combined writer ───────────────────────────────────────────────────


def write_diff_reports(diff: WorkbookDiff) -> tuple[Path, Path]:
    """Convenience: write both Markdown and CSV reports for a diff."""
    md = write_markdown_report(diff) if config.GENERATE_MARKDOWN_REPORT else None
    csv_p = write_csv_report(diff) if config.GENERATE_CSV_REPORT else None
    return md, csv_p


# ── Rollup reports across all diffs ───────────────────────────────────


# ── Block timeline (single-block history across all months) ──────────


def write_block_timeline_report(
    sheet_name: str,
    primary_labels: str | list[str],
    *,
    limit: int | None = None,
    months: int | None = None,
    output_path: Optional[Path] = None,
) -> tuple[Path, Path]:
    """
    Write a per-block timeline report covering all recorded changes
    for a given (sheet, block(s)) pair across the database.

    Args:
        sheet_name: Sheet to filter on (exact match).
        primary_labels: One or more block primary labels to filter on.
        limit: If set, include only the most recent N changes.
        months: If set, include only changes from the last N months.
        output_path: Markdown destination. CSV gets a sibling .csv file.
            Defaults to reports/timeline_<sheet>_<block>.md.

    Returns:
        (md_path, csv_path) tuple.
    """
    from .storage import find_block_timeline

    if isinstance(primary_labels, str):
        primary_labels = [primary_labels]
    multi = len(primary_labels) > 1

    rows = find_block_timeline(sheet_name, primary_labels, limit=limit, months=months)

    safe_sheet = "".join(c if c.isalnum() or c in "-_" else "_" for c in sheet_name)
    if output_path is None:
        from datetime import datetime
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = config.REPORT_DIR / f"timeline_{safe_sheet}_{ts}.md"
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    csv_path = output_path.with_suffix(".csv")

    # Markdown
    lines: list[str] = []
    if multi:
        labels_joined = ", ".join(f"`{lbl}`" for lbl in primary_labels)
        lines.append(f"# Block Timeline: {labels_joined} on sheet `{sheet_name}`")
    else:
        lines.append(f"# Block Timeline: `{primary_labels[0]}` on sheet `{sheet_name}`")
    lines.append("")
    if limit:
        lines.append(f"_Showing the last {limit} changes._")
    if months:
        lines.append(f"_Showing changes from the last {months} months._")
    lines.append(f"**Total changes recorded:** {len(rows)}")
    lines.append("")

    if not rows:
        lines.append("_No changes found for this block._")
        lines.append("")
        lines.append("Possible reasons: the block name is misspelled, the sheet "
                     "name is wrong, or the block has been stable across all tracked "
                     "months. Use `excel-evo-tracker history \"<label>\"` to search "
                     "for the element across sheets.")
    else:
        lines.append("## Timeline")
        lines.append("")
        if multi:
            lines.append("| Block | Old File | New File | Change | Confidence | Review | Detail |")
            lines.append("|-------|----------|----------|--------|-----------:|:------:|--------|")
            for r in rows:
                old_lbl = _display_label(r["old_file"], r["old_month"])
                new_lbl = _display_label(r["new_file"], r["new_month"])
                conf = f"{r['confidence']:.2f}" if r["confidence"] else "—"
                review = "⚠️" if r["needs_review"] else ""
                detail = (r["detail"] or "").replace("|", "\\|").replace("\n", " ")
                lines.append(
                    f"| {r['element_name']} | {old_lbl} | {new_lbl} | {r['change_type']} | {conf} | {review} | {detail} |"
                )
        else:
            lines.append("| Old File | New File | Change | Confidence | Review | Detail |")
            lines.append("|----------|----------|--------|-----------:|:------:|--------|")
            for r in rows:
                old_lbl = _display_label(r["old_file"], r["old_month"])
                new_lbl = _display_label(r["new_file"], r["new_month"])
                conf = f"{r['confidence']:.2f}" if r["confidence"] else "—"
                review = "⚠️" if r["needs_review"] else ""
                detail = (r["detail"] or "").replace("|", "\\|").replace("\n", " ")
                lines.append(
                    f"| {old_lbl} | {new_lbl} | {r['change_type']} | {conf} | {review} | {detail} |"
                )
        lines.append("")

        # Aggregate by change type
        type_counts: dict[str, int] = {}
        for r in rows:
            t = r["change_type"]
            type_counts[t] = type_counts.get(t, 0) + 1
        lines.append("## Summary by Change Type")
        lines.append("")
        lines.append("| Change Type | Count |")
        lines.append("|-------------|------:|")
        for t in sorted(type_counts):
            lines.append(f"| {t} | {type_counts[t]} |")
        lines.append("")

    output_path.write_text("\n".join(lines), encoding="utf-8")

    # CSV
    csv_fields = (
        ["block", "old_month", "new_month", "change_type", "confidence",
         "needs_review", "impact_weight", "old_file", "new_file", "detail"]
        if multi else [
            "old_month", "new_month", "change_type", "confidence",
            "needs_review", "impact_weight", "old_file", "new_file", "detail",
        ]
    )
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=csv_fields, extrasaction="ignore")
        writer.writeheader()
        for r in rows:
            row = {
                "block": r["element_name"] if multi else None,
                "old_month": r["old_month"] or "",
                "new_month": r["new_month"] or "",
                "change_type": r["change_type"],
                "confidence": r["confidence"] if r["confidence"] is not None else "",
                "needs_review": "Y" if r["needs_review"] else "",
                "impact_weight": r["impact_weight"] if r["impact_weight"] is not None else "",
                "old_file": r["old_file"],
                "new_file": r["new_file"],
                "detail": r["detail"] or "",
            }
            writer.writerow(row)

    logger.info("Wrote block timeline: %s (%d changes)", output_path, len(rows))
    return output_path, csv_path


def write_rollup_reports(output_dir: Optional[Path] = None) -> tuple[Path, Path]:
    """
    Write a rollup index covering every diff in the database.

    Produces two files in the report directory:
      - rollup.md  — sorted by impact score, links to per-diff reports
      - rollup.csv — same data as a flat table for filtering in Excel
    """
    output_dir = Path(output_dir) if output_dir else config.REPORT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    all_diffs = list_diffs()
    by_impact = sorted(all_diffs, key=lambda d: -d["impact_score"])

    # Markdown rollup
    md_path = output_dir / "rollup.md"
    lines: list[str] = []
    lines.append("# Excel Evolution Tracker — Rollup Report")
    lines.append("")
    lines.append(f"Total diffs: **{len(all_diffs)}**")
    lines.append("")
    lines.append("## Diffs by Impact (descending)")
    lines.append("")
    if not by_impact:
        lines.append("_No diffs recorded yet._")
    else:
        lines.append("| Old File | New File | Impact | Changes | Review | Report |")
        lines.append("|----------|----------|-------:|--------:|-------:|--------|")
        for d in by_impact:
            old = _display_label(d["old_file"], d["old_month"])
            new = _display_label(d["new_file"], d["new_month"])
            md_link = ""
            if d["md_report_path"]:
                md_link = f"[md]({Path(d['md_report_path']).name})"
            bar = _impact_bar(d["impact_score"], width=10)
            lines.append(
                f"| {old} | {new} | {d['impact_score']:.3f} `{bar}` | "
                f"{d['total_changes']} | {d['needs_review_count']} | {md_link} |"
            )
    lines.append("")
    md_path.write_text("\n".join(lines), encoding="utf-8")

    # CSV rollup
    csv_path = output_dir / "rollup.csv"
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "old_month", "new_month", "old_file", "new_file",
                "diff_date", "impact_score", "total_changes",
                "needs_review_count", "json_path", "md_report_path",
            ],
        )
        writer.writeheader()
        for d in all_diffs:
            writer.writerow({
                "old_month": d["old_month"] or "",
                "new_month": d["new_month"] or "",
                "old_file": d["old_file"],
                "new_file": d["new_file"],
                "diff_date": d["diff_date"],
                "impact_score": d["impact_score"],
                "total_changes": d["total_changes"],
                "needs_review_count": d["needs_review_count"],
                "json_path": d["json_path"],
                "md_report_path": d["md_report_path"] or "",
            })

    logger.info("Wrote rollup reports: %s, %s", md_path, csv_path)
    return md_path, csv_path
