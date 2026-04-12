"""
Data models for the Excel Evolution Tracker.

All models are Pydantic BaseModels so they serialize cleanly to JSON and
validate on construction. The hierarchy mirrors the domain:

    WorkbookSnapshot
    ├── SheetSnapshot
    │   ├── Cell (sparse map, keyed by address)
    │   └── Block (detected clusters)
    └── NamedRange

    WorkbookDiff
    ├── SheetChange
    ├── BlockChange
    ├── NamedRangeChange
    └── MergedRegionChange
"""

from __future__ import annotations

from enum import Enum
from typing import Optional, Union

from pydantic import BaseModel, Field


# ── Enums ─────────────────────────────────────────────────────────────


class CellType(str, Enum):
    LABEL = "label"          # Text string — the primary tracking unit
    NUMERIC = "numeric"      # Number — used for shape but not identity
    BOOLEAN = "boolean"
    DATE = "date"
    EMPTY = "empty"          # Should never be stored, but defined for completeness


class ChangeType(str, Enum):
    ADDED = "added"
    REMOVED = "removed"
    MOVED = "moved"
    RENAMED = "renamed"
    RESHAPED = "reshaped"
    LABEL_CHANGED = "label_changed"
    MODIFIED = "modified"       # Internal change within a matched element
    REORDERED = "reordered"     # Same sheets, different order


# ── Cell & Sheet ──────────────────────────────────────────────────────


class Cell(BaseModel):
    """A single non-empty cell in a worksheet."""
    address: str                               # e.g. "B5"
    row: int                                   # 1-indexed (Excel convention)
    col: int                                   # 1-indexed (A=1, B=2, ...)
    value: Optional[Union[str, float, int, bool]] = None
    dtype: CellType
    is_merged: bool = False
    merge_range: Optional[str] = None          # e.g. "C7:E7" if this cell is in a merge


class Block(BaseModel):
    """
    A detected cluster of related cells within a sheet.

    Blocks are discovered via connected-component analysis on the cell
    occupancy grid (with gap tolerance). The primary_label is the best
    candidate for a human-facing identity.
    """
    block_id: str                              # Unique within snapshot, e.g. "Summary_B3_R3C2"
    sheet_name: str
    top_left: tuple[int, int]                  # (row, col) — 1-indexed
    bottom_right: tuple[int, int]
    shape: tuple[int, int]                     # (row_count, col_count)
    cell_count: int                            # Total non-empty cells in block

    labels: list[str]                          # All text values found in the block
    primary_label: Optional[str] = None        # Best identity candidate

    label_cells: list[Cell] = Field(default_factory=list)
    value_cells: list[Cell] = Field(default_factory=list)

    fingerprint: str                           # Position-independent hash of internal structure
    neighbor_context: list[str] = Field(default_factory=list)  # Nearby blocks' primary labels


class SheetSnapshot(BaseModel):
    """Complete structural snapshot of a single worksheet."""
    name: str
    index: int                                 # Position in workbook (0-indexed)
    is_hidden: bool = False

    used_range: Optional[str] = None           # e.g. "A1:Z100"
    max_row: int = 0
    max_col: int = 0

    # Sparse cell map: only non-empty cells, keyed by address
    cells: dict[str, Cell] = Field(default_factory=dict)

    # Detected blocks (populated by block_detector)
    blocks: list[Block] = Field(default_factory=list)

    # Merged cell ranges as A1-style strings, e.g. ["A1:D1", "B5:B10"]
    merged_regions: list[str] = Field(default_factory=list)


# ── Named range ───────────────────────────────────────────────────────


class NamedRange(BaseModel):
    name: str
    scope: str                                 # "workbook" or a sheet name
    refers_to: str                             # e.g. "Summary!$C$4:$C$5"


# ── Workbook snapshot ─────────────────────────────────────────────────


class WorkbookSnapshot(BaseModel):
    """Top-level snapshot of an entire workbook at a point in time."""
    file_name: str
    file_path: str
    file_hash: str                             # MD5 of the original XLSB
    snapshot_date: str                         # ISO timestamp of extraction
    month_label: Optional[str] = None          # User-assigned identifier, e.g. "2024-01"

    sheet_names: list[str] = Field(default_factory=list)   # Ordered
    sheets: dict[str, SheetSnapshot] = Field(default_factory=dict)
    named_ranges: list[NamedRange] = Field(default_factory=list)

    total_blocks: int = 0
    total_cells: int = 0
    metadata: dict = Field(default_factory=dict)


# ── Change records ────────────────────────────────────────────────────


class SheetChange(BaseModel):
    change_type: ChangeType
    sheet_name: str                            # Canonical name to report
    old_name: Optional[str] = None
    new_name: Optional[str] = None
    old_index: Optional[int] = None
    new_index: Optional[int] = None
    detail: str = ""


class BlockChange(BaseModel):
    change_type: ChangeType
    sheet_name: str

    # For added: only new_block is set; for removed: only old_block
    old_block: Optional[Block] = None
    new_block: Optional[Block] = None

    # Match confidence when both blocks exist (moved/reshaped/label_changed/modified)
    match_confidence: float = 0.0
    needs_review: bool = False

    # Optional deltas depending on change_type
    position_delta: Optional[tuple[int, int]] = None     # (row_shift, col_shift)
    label_diff: Optional[dict] = None                    # {"old": "...", "new": "..."}
    shape_diff: Optional[dict] = None                    # {"old": (r,c), "new": (r,c)}
    internal_labels_added: list[str] = Field(default_factory=list)
    internal_labels_removed: list[str] = Field(default_factory=list)

    detail: str = ""


class NamedRangeChange(BaseModel):
    change_type: ChangeType
    name: str
    old_refers_to: Optional[str] = None
    new_refers_to: Optional[str] = None
    old_scope: Optional[str] = None
    new_scope: Optional[str] = None
    detail: str = ""


class MergedRegionChange(BaseModel):
    change_type: ChangeType                    # ADDED or REMOVED
    sheet_name: str
    region: str                                # e.g. "A1:D1"


# ── Workbook diff ─────────────────────────────────────────────────────


class DiffSummary(BaseModel):
    """Counts of each change category — populated by the differ."""
    sheets_added: int = 0
    sheets_removed: int = 0
    sheets_renamed: int = 0
    sheets_reordered: int = 0

    blocks_added: int = 0
    blocks_removed: int = 0
    blocks_moved: int = 0
    blocks_reshaped: int = 0
    blocks_label_changed: int = 0
    blocks_modified: int = 0

    named_ranges_added: int = 0
    named_ranges_removed: int = 0
    named_ranges_modified: int = 0

    merges_added: int = 0
    merges_removed: int = 0

    needs_review_count: int = 0
    total_changes: int = 0


class WorkbookDiff(BaseModel):
    """Complete structural diff between two workbook snapshots."""
    old_file: str
    new_file: str
    old_month: Optional[str] = None
    new_month: Optional[str] = None
    diff_date: str                             # ISO timestamp of the diff run

    sheet_changes: list[SheetChange] = Field(default_factory=list)
    block_changes: list[BlockChange] = Field(default_factory=list)
    named_range_changes: list[NamedRangeChange] = Field(default_factory=list)
    merged_region_changes: list[MergedRegionChange] = Field(default_factory=list)

    summary: DiffSummary = Field(default_factory=DiffSummary)
    impact_score: float = 0.0                  # Normalized to [0, 1]
