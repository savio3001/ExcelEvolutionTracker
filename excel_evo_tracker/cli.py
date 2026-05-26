"""
Command-line interface for the Excel Evolution Tracker.

Subcommands:

  batch        Process every XLSB in a directory; build full history.
  incremental  Process one new XLSB and diff against the latest stored.
  compare      Diff two XLSB files directly without touching the DB.
  rollup       Regenerate the cross-diff rollup index.
  history      Show recorded changes for a specific element across all months.
  review       List all changes flagged needs_review.

Run `python -m excel_evo_tracker.cli <command> --help` for details.
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from . import config
from .pipeline import run_batch, run_compare, run_incremental
from .reporter import write_rollup_reports
from .storage import find_changes_for_element, find_changes_needing_review, init_db


def _setup_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s  %(levelname)-7s  %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


# ── Subcommand handlers ───────────────────────────────────────────────


def cmd_batch(args: argparse.Namespace) -> int:
    result = run_batch(
        xlsb_dir=Path(args.input) if args.input else None,
        pattern=args.pattern,
        write_block_debug=args.block_debug,
        write_rollup=not args.no_rollup,
        reset_db=args.reset_db,
    )
    print()
    print(f"Processed {result['snapshots']} snapshots, generated {result['diffs']} diffs.")
    print(f"Reports written to: {result['report_dir']}")
    if result["errors"]:
        print(f"\n{len(result['errors'])} error(s):")
        for e in result["errors"]:
            print(f"  - {e}")
        return 1
    return 0


def cmd_incremental(args: argparse.Namespace) -> int:
    result = run_incremental(
        new_xlsb=Path(args.file),
        month_label=args.month,
        write_block_debug=args.block_debug,
        write_rollup=not args.no_rollup,
    )
    print()
    if result.get("is_first"):
        print(f"First file in database: {result['snapshot']}")
        print("No diff generated. Add another file to start tracking changes.")
    else:
        print(f"Compared {result['previous']} → {result['snapshot']}")
        print(f"  Impact score: {result['impact_score']:.3f}")
        print(f"  Total changes: {result['total_changes']}")
        print(f"  Needs review: {result['needs_review']}")
        print(f"  Markdown report: {result['diff']}")
        print(f"  CSV report:      {result['csv']}")
    return 0


def cmd_replace(args: argparse.Namespace) -> int:
    """Replace an existing snapshot with a corrected version."""
    from .pipeline import run_replace
    from .storage import get_connection

    init_db(force=False)

    # Validate the old file exists in DB
    with get_connection() as conn:
        exists = conn.execute(
            "SELECT 1 FROM snapshots WHERE file_name = ?", (args.old,)
        ).fetchone()
    if not exists:
        print(f"No snapshot found for file {args.old!r}.")
        print("Available files:")
        with get_connection() as conn:
            files = conn.execute(
                "SELECT DISTINCT file_name FROM snapshots ORDER BY file_name"
            ).fetchall()
        for f in files[:30]:
            print(f"  | {f['file_name']}")
        return 1

    new_path = Path(args.new)
    if not new_path.exists():
        print(f"Replacement file not found: {new_path}")
        return 1

    print(f"Replacing {args.old!r} with {new_path.name}")
    result = run_replace(
        old_file_name=args.old,
        new_xlsb=new_path,
        month_label=args.month,
    )

    if "error" in result:
        print(f"Error: {result['error']}")
        return 1

    purged = result["purged"]
    print(f"\nPurged:")
    print(f"  Snapshot rows:  {purged['snapshot_rows']}")
    print(f"  Diff rows:      {purged['diff_rows']}")
    print(f"  Change rows:    {purged['change_rows']}")
    print(f"  Files deleted:  {len(purged['files'])}")
    print(f"\nRegenerated:")
    print(f"  New file:       {result['new_file']}")
    print(f"  Diffs created:  {result['diffs_regenerated']}")
    if result["prev_neighbor"]:
        print(f"  Prev neighbor:  {result['prev_neighbor']} → {result['new_file']}")
    if result["next_neighbor"]:
        print(f"  Next neighbor:  {result['new_file']} → {result['next_neighbor']}")
    return 0


def cmd_compare(args: argparse.Namespace) -> int:
    diff = run_compare(
        old_xlsb=Path(args.old),
        new_xlsb=Path(args.new),
        write_reports=True,
    )
    print()
    print(f"Comparison: {diff.old_file} → {diff.new_file}")
    print(f"  Impact score: {diff.impact_score:.3f}")
    print(f"  Total changes: {diff.summary.total_changes}")
    print(f"  Needs review: {diff.summary.needs_review_count}")
    print(f"  Reports written to: {config.REPORT_DIR}")
    return 0


def cmd_rollup(args: argparse.Namespace) -> int:
    init_db(force=False)
    md, csv_p = write_rollup_reports()
    print(f"Rollup written: {md}")
    print(f"               {csv_p}")
    return 0


def cmd_history(args: argparse.Namespace) -> int:
    init_db(force=False)
    rows = find_changes_for_element(args.element)
    if not rows:
        print(f"No changes recorded for element: {args.element!r}")
        return 0
    print(f"History for element {args.element!r}: {len(rows)} change(s)")
    print()
    for r in rows:
        old = Path(r["old_file"]).stem if r.get("old_file") else (r["old_month"] or "?")
        new = Path(r["new_file"]).stem if r.get("new_file") else (r["new_month"] or "?")
        print(f"  [{old} → {new}] {r['change_category']}/{r['change_type']}: {r['detail']}")
    return 0


def cmd_review(args: argparse.Namespace) -> int:
    init_db(force=False)
    rows = find_changes_needing_review()
    if not rows:
        print("No changes flagged for review.")
        return 0
    print(f"{len(rows)} change(s) flagged for review:")
    print()
    for r in rows:
        old = Path(r["old_file"]).stem if r.get("old_file") else (r["old_month"] or "?")
        new = Path(r["new_file"]).stem if r.get("new_file") else (r["new_month"] or "?")
        conf = f" (conf {r['confidence']:.2f})" if r["confidence"] else ""
        print(f"  [{old} → {new}] {r['change_category']}/{r['change_type']}{conf}")
        print(f"    element: {r['element_name']}")
        print(f"    detail:  {r['detail']}")
        print()
    return 0


def _col_letters_to_num(letters: str) -> int:
    """'A' → 1, 'Z' → 26, 'AA' → 27, 'CS' → 97."""
    n = 0
    for c in letters.upper():
        if not c.isalpha():
            raise ValueError(f"Invalid column letters: {letters!r}")
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n


def _parse_range_filter(range_str: str) -> tuple[int | None, int | None, int | None, int | None]:
    """
    Parse a range string into (min_row, min_col, max_row, max_col).

    None on any axis means "unbounded on this side." Supported forms:

      "CS1:DA40"  → rows 1-40,  cols CS-DA
      "CS:DA"     → all rows,   cols CS-DA
      "1:40"      → rows 1-40,  all cols
      "CS1"       → single cell at row 1, col CS
      "5"         → single row 5
      "CS"        → single column CS

    Raises ValueError on malformed input.
    """
    import re
    s = range_str.strip()
    if not s:
        raise ValueError("Empty range")

    # Split on ':' if present
    if ":" in s:
        left, right = s.split(":", 1)
        left, right = left.strip(), right.strip()
    else:
        left = right = s

    # A token is either: pure digits (row), pure letters (col), or letters+digits (cell)
    cell_re = re.compile(r"^([A-Za-z]*)(\d*)$")

    def parse_token(tok: str) -> tuple[int | None, int | None]:
        m = cell_re.match(tok)
        if not m or not tok:
            raise ValueError(f"Invalid range token: {tok!r}")
        letters, digits = m.group(1), m.group(2)
        col = _col_letters_to_num(letters) if letters else None
        row = int(digits) if digits else None
        if col is None and row is None:
            raise ValueError(f"Invalid range token: {tok!r}")
        return row, col

    left_row, left_col = parse_token(left)
    right_row, right_col = parse_token(right)

    # Determine bounds — if one side is missing a row/col, copy from the other
    min_row = left_row if left_row is not None else right_row
    max_row = right_row if right_row is not None else left_row
    min_col = left_col if left_col is not None else right_col
    max_col = right_col if right_col is not None else left_col

    # If both sides are pure-row (e.g. "1:40"), cols stay None → unbounded
    # If both sides are pure-col (e.g. "CS:DA"), rows stay None → unbounded

    # Order normalization
    if min_row is not None and max_row is not None and min_row > max_row:
        min_row, max_row = max_row, min_row
    if min_col is not None and max_col is not None and min_col > max_col:
        min_col, max_col = max_col, min_col

    return min_row, min_col, max_row, max_col


def _block_intersects_range(
    block,
    min_row: int | None,
    min_col: int | None,
    max_row: int | None,
    max_col: int | None,
) -> bool:
    """Check whether a block's bounding box overlaps the given range."""
    bt, bl = block.top_left
    bb, br = block.bottom_right
    # Row overlap (or unbounded)
    if min_row is not None and bb < min_row:
        return False
    if max_row is not None and bt > max_row:
        return False
    # Col overlap (or unbounded)
    if min_col is not None and br < min_col:
        return False
    if max_col is not None and bl > max_col:
        return False
    return True


def cmd_blocks(args: argparse.Namespace) -> int:
    """List blocks detected on a sheet from a stored snapshot."""
    from .storage import get_latest_snapshot, get_connection, load_snapshot
    from pathlib import Path

    init_db(force=False)

    # Pick which snapshot to read
    snap = None
    if args.file:
        # Look up snapshot by file_name (most recent matching row wins)
        with get_connection() as conn:
            row = conn.execute(
                "SELECT json_path FROM snapshots WHERE file_name = ? "
                "ORDER BY snapshot_date DESC LIMIT 1",
                (args.file,),
            ).fetchone()
        if row is None:
            print(f"No snapshot found for file {args.file!r}.")
            print("Available files:")
            with get_connection() as conn:
                files = conn.execute(
                    "SELECT DISTINCT file_name FROM snapshots ORDER BY file_name"
                ).fetchall()
            for f in files[:30]:
                print(f"  | {f['file_name']}")
            return 1
        snap = load_snapshot(Path(row["json_path"]))
    else:
        snap = get_latest_snapshot()
        if snap is None:
            print("No snapshots in database. Run `batch` or `incremental` first.")
            return 1

    if args.sheet not in snap.sheets:
        print(f"Sheet {args.sheet!r} not found in snapshot {snap.file_name!r}.")
        print("Available sheets:")
        for s in snap.sheet_names[:50]:
            print(f"  | {s}")
        if len(snap.sheet_names) > 50:
            print(f"  ... ({len(snap.sheet_names) - 50} more)")
        return 1

    sheet = snap.sheets[args.sheet]
    print(f"Snapshot:  {snap.file_name}")
    print(f"Sheet:     {args.sheet}  ({sheet.max_row}×{sheet.max_col}, "
          f"{len(sheet.cells)} cells, {len(sheet.blocks)} blocks)")
    print()

    if not sheet.blocks:
        print("(No blocks detected on this sheet — possibly in SKIP_BLOCK_DETECTION_SHEETS, "
              "or block detection produced no clusters.)")
        return 0

    # Optional filter by primary label substring
    blocks = sheet.blocks
    if args.filter:
        needle = args.filter.lower()
        blocks = [b for b in blocks if b.primary_label and needle in b.primary_label.lower()]

    # Optional filter by range
    range_summary = ""
    if args.range:
        try:
            min_row, min_col, max_row, max_col = _parse_range_filter(args.range)
        except ValueError as e:
            print(f"Invalid range {args.range!r}: {e}")
            print("Examples: 'CS1:DA40' (cells), 'CS:DA' (columns), '1:40' (rows)")
            return 1
        blocks = [
            b for b in blocks
            if _block_intersects_range(b, min_row, min_col, max_row, max_col)
        ]
        range_summary = f"  intersecting {args.range!r}"

    if not blocks:
        what = []
        if args.filter:
            what.append(f"label filter {args.filter!r}")
        if args.range:
            what.append(f"range {args.range!r}")
        if what:
            print(f"No blocks matching {' and '.join(what)}.")
        else:
            print("(No blocks detected on this sheet.)")
        return 0

    # Build the address range string for each block
    def addr(pos):
        row, col = pos
        letters = ""
        c = col
        while c > 0:
            c, rem = divmod(c - 1, 26)
            letters = chr(ord("A") + rem) + letters
        return f"{letters}{row}"

    # Print as a table
    print(f"  {'#':>3}  {'primary_label':<120}  {'range':<14}  {'shape':<10}  cells  labels")
    print(f"  {'-'*3}  {'-'*120}  {'-'*14}  {'-'*10}  {'-'*5}  {'-'*6}")
    for i, b in enumerate(blocks, start=1):
        rng = f"{addr(b.top_left)}:{addr(b.bottom_right)}"
        shape = f"{b.shape[0]}×{b.shape[1]}"
        label = b.primary_label or "(no label)"
        print(f"  {i:>3}  {label:<120}  {rng:<14}  {shape:<10}  {b.cell_count:>5}  {len(b.label_cells):>6}")

    print()
    print(f"Total: {len(blocks)} blocks shown{range_summary}")
    if args.filter or args.range:
        applied = []
        if args.filter:
            applied.append(f"label {args.filter!r}")
        if args.range:
            applied.append(f"range {args.range!r}")
        print(f"  (filtered from {len(sheet.blocks)} total by {' + '.join(applied)})")
    return 0


def cmd_timeline(args: argparse.Namespace) -> int:
    """Generate a timeline report for specific block(s) on a specific sheet."""
    from .reporter import write_block_timeline_report
    from .storage import list_known_block_labels
    init_db(force=False)

    sheet_name = args.sheet
    block_labels = args.block

    # Optional fuzzy matching against known block labels for the sheet
    if args.fuzzy:
        from rapidfuzz import fuzz
        known = list_known_block_labels(sheet_name)
        if not known:
            print(f"No blocks recorded for sheet {sheet_name!r}.")
            return 1
        matched = []
        for lbl in block_labels:
            scored = sorted(
                ((s, fuzz.ratio(lbl.lower(), s.lower())) for s in known),
                key=lambda t: -t[1],
            )
            if scored[0][1] < 60:
                print(f"No fuzzy match for block {lbl!r} on sheet {sheet_name!r}.")
                print(f"Best candidates:")
                for s, score in scored[:5]:
                    print(f"  {score:3d}  {s}")
                return 1
            matched.append(scored[0][0])
            print(f"Fuzzy matched {lbl!r} → {scored[0][0]!r} (score {scored[0][1]})")
        block_labels = matched

    out_path = Path(args.out) if args.out else None
    md_path, csv_path = write_block_timeline_report(
        sheet_name=sheet_name,
        primary_labels=block_labels,
        limit=args.limit,
        months=args.months,
        output_path=out_path,
    )
    print(f"Timeline report written:")
    print(f"  Markdown: {md_path}")
    print(f"  CSV:      {csv_path}")
    return 0


def cmd_stability(args: argparse.Namespace) -> int:
    """Audit the learned cell stability classifier."""
    from .stability import (
        CellRole, dump_sheet_classification, list_known_sheets,
        ensure_stability_schema,
    )
    init_db(force=False)
    ensure_stability_schema()

    # Collect requested sheet names: positional first, then comma-separated
    requested: list[str] = list(args.sheets or [])
    if args.sheet_list:
        requested.extend(s.strip() for s in args.sheet_list.split(",") if s.strip())

    known = list_known_sheets()
    if not known:
        print("Ledger is empty. Run `batch` or `incremental` first to populate it.")
        return 0

    # Resolve which sheets to dump
    if not requested:
        targets = known
    elif args.fuzzy:
        from rapidfuzz import fuzz
        targets = []
        for req in requested:
            matches = [(s, fuzz.ratio(req.lower(), s.lower())) for s in known]
            matches.sort(key=lambda t: -t[1])
            picks = [s for s, score in matches if score >= 70]
            if picks:
                targets.extend(picks)
            else:
                print(f"No fuzzy match for {req!r} (best: {matches[0] if matches else 'none'})")
        targets = list(dict.fromkeys(targets))   # de-dup preserving order
    else:
        targets = []
        for req in requested:
            if req in known:
                targets.append(req)
            else:
                print(f"Sheet {req!r} not in ledger. Use --fuzzy or try one of:")
                for s in known[:20]:
                    print(f"  | {s}")
                return 1

    for sheet_name in targets:
        rows = dump_sheet_classification(sheet_name)
        print()
        print(f"=== Sheet: {sheet_name}  ({len(rows)} tracked cells) ===")
        if not rows:
            print("  (no cells tracked)")
            continue
        print(f"  {'primary_label':<40}  {'pos':<8}  {'role':<8}  {'score':<6}  {'obs':<4}  last_value")
        print(f"  {'-'*40}  {'-'*8}  {'-'*8}  {'-'*6}  {'-'*4}  {'-'*30}")
        for r in rows:
            pos = f"R{r['rel_row']}C{r['rel_col']}"
            label = (r["primary_label"] or "")[:38]
            last = (r["last_value"] or "")[:30]
            print(f"  {label:<40}  {pos:<8}  {r['role']:<8}  "
                  f"{r['stability_score']:<6.2f}  {r['observations']:<4}  {last}")
    return 0


# ── Argument parser ───────────────────────────────────────────────────


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="excel-evo-tracker",
        description="Detect structural drift across monthly XLSB templates.",
    )
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable debug logging")

    subs = parser.add_subparsers(dest="command", required=True, metavar="COMMAND")

    # batch
    p_batch = subs.add_parser("batch", help="Process every XLSB in a directory")
    p_batch.add_argument("--input", "-i", help="Source directory (default: config.XLSB_INPUT_DIR)")
    p_batch.add_argument("--pattern", default="*.xlsb", help="Glob pattern (default: *.xlsb)")
    p_batch.add_argument("--block-debug", action="store_true",
                         help="Write per-file block detection debug reports")
    p_batch.add_argument("--no-rollup", action="store_true", help="Skip rollup report")
    p_batch.add_argument("--reset-db", action="store_true",
                         help="Drop and recreate the database before processing")
    p_batch.set_defaults(func=cmd_batch)

    # incremental
    p_inc = subs.add_parser("incremental", help="Process one new XLSB against latest stored")
    p_inc.add_argument("file", help="Path to the new XLSB file")
    p_inc.add_argument("--month", "-m", help="Month label (e.g. 2024-12); inferred from filename if omitted")
    p_inc.add_argument("--block-debug", action="store_true",
                       help="Write a block detection debug report")
    p_inc.add_argument("--no-rollup", action="store_true", help="Skip rollup refresh")
    p_inc.set_defaults(func=cmd_incremental)

    # replace
    p_rep = subs.add_parser("replace",
                            help="Replace an existing snapshot with a corrected version of the file")
    p_rep.add_argument("--old", required=True,
                       help="File name of the snapshot to replace (as shown in `rollup`)")
    p_rep.add_argument("--new", required=True,
                       help="Path to the replacement XLSB file")
    p_rep.add_argument("--month", "-m",
                       help="Month label for the replacement (optional)")
    p_rep.set_defaults(func=cmd_replace)

    # compare
    p_cmp = subs.add_parser("compare", help="Diff two XLSB files without touching the DB")
    p_cmp.add_argument("--old", required=True, help="Path to the older XLSB")
    p_cmp.add_argument("--new", required=True, help="Path to the newer XLSB")
    p_cmp.set_defaults(func=cmd_compare)

    # rollup
    p_roll = subs.add_parser("rollup", help="Regenerate the cross-diff rollup index")
    p_roll.set_defaults(func=cmd_rollup)

    # history
    p_hist = subs.add_parser("history", help="Show change history for a named element")
    p_hist.add_argument("element", help="Element name (block primary label, sheet name, or named range)")
    p_hist.set_defaults(func=cmd_history)

    # review
    p_rev = subs.add_parser("review", help="List all changes flagged needs_review")
    p_rev.set_defaults(func=cmd_review)

    # stability
    p_stab = subs.add_parser("stability", help="Audit the learned cell classifier")
    p_stab.add_argument("sheets", nargs="*",
                        help="Sheet name(s) to inspect. Quote names containing spaces. "
                             "Omit to dump all sheets.")
    p_stab.add_argument("--sheets", dest="sheet_list",
                        help="Comma-separated sheet names (alternative to positional)")
    p_stab.add_argument("--fuzzy", action="store_true",
                        help="Use fuzzy matching on sheet names")
    p_stab.set_defaults(func=cmd_stability)

    # timeline
    p_tl = subs.add_parser("timeline",
                           help="Generate a timeline report for one or more blocks on a sheet")
    p_tl.add_argument("sheet", help="Sheet name (quote if it contains spaces)")
    p_tl.add_argument("block", nargs="+",
                      help="One or more block primary labels (quote if they contain spaces)")
    p_tl.add_argument("--limit", "-n", type=int, default=None,
                      help="Show only the most recent N changes (default: all)")
    p_tl.add_argument("--months", "-m", type=int, default=None,
                      help="Show only changes from the last N months (default: all)")
    p_tl.add_argument("--fuzzy", action="store_true",
                      help="Fuzzy-match the block label against known labels for the sheet")
    p_tl.add_argument("--out", "-o", help="Output Markdown path (CSV gets a sibling .csv)")
    p_tl.set_defaults(func=cmd_timeline)

    # blocks
    p_blk = subs.add_parser("blocks",
                            help="List all blocks detected on a sheet from a stored snapshot")
    p_blk.add_argument("sheet", help="Sheet name (quote if it contains spaces)")
    p_blk.add_argument("range", nargs="?", default=None,
                       help="Optional range filter. Examples: 'CS1:DA40' (cell range), "
                            "'CS:DA' (entire columns), '1:40' (entire rows)")
    p_blk.add_argument("--file", "-f",
                       help="File name to read blocks from (default: latest snapshot)")
    p_blk.add_argument("--filter",
                       help="Substring filter on block primary label (case-insensitive)")
    p_blk.set_defaults(func=cmd_blocks)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    _setup_logging(args.verbose)
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
