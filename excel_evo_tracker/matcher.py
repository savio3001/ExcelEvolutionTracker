"""
Multi-pass block matcher.

Given two lists of blocks (old and new) from a matched sheet pair,
produces a list of MatchPair objects identifying which blocks correspond
to each other across the snapshots, plus the unmatched leftovers
(which the differ will report as ADDED / REMOVED).

Strategy (in order):

  Pass 1 — Exact fingerprint match
    Blocks whose fingerprints are identical and unique on both sides
    are paired immediately at high confidence (0.95). If a fingerprint
    has multiple candidates, position proximity breaks the tie.

  Pass 2 — Hybrid composite score
    Remaining blocks are scored against each other using a weighted
    combination of:
        - primary label similarity (rapidfuzz)
        - fingerprint match (binary 0/1)
        - position proximity (Manhattan distance, normalized)
        - neighbor context overlap (Jaccard)
        - shape similarity
    Weights live in config.MATCH_WEIGHTS. Greedy assignment in
    descending score order; pairs above AUTO_ACCEPT_THRESHOLD are
    accepted, pairs in [REVIEW_ZONE_THRESHOLD, AUTO_ACCEPT) are
    accepted with needs_review=True.

  Pass 3 — Label set Jaccard fallback
    For blocks that have very different primary labels but share most
    of their internal label set (a renamed-but-otherwise-stable block),
    Jaccard overlap on label sets catches them. Lower confidence cap.

Anything still unmatched after Pass 3 is left for the differ to report
as added or removed.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field

from rapidfuzz import fuzz

from . import config
from .models import Block

logger = logging.getLogger(__name__)


# ── Result types ──────────────────────────────────────────────────────


@dataclass
class MatchPair:
    """A matched (old, new) block pair with confidence and provenance."""
    old_block: Block
    new_block: Block
    confidence: float
    pass_used: int                          # 1, 2, or 3
    needs_review: bool = False
    score_breakdown: dict[str, float] = field(default_factory=dict)


@dataclass
class MatchResult:
    """Output of running the matcher on two block lists."""
    pairs: list[MatchPair] = field(default_factory=list)
    unmatched_old: list[Block] = field(default_factory=list)
    unmatched_new: list[Block] = field(default_factory=list)


# ── Scoring helpers ───────────────────────────────────────────────────


def _label_similarity(a: Block, b: Block) -> float:
    """Fuzzy ratio between primary labels, normalized to [0, 1]."""
    if not a.primary_label or not b.primary_label:
        return 0.0
    return fuzz.ratio(a.primary_label, b.primary_label) / 100.0


def _fingerprint_match(a: Block, b: Block) -> float:
    """1.0 if fingerprints identical, else 0.0."""
    return 1.0 if a.fingerprint == b.fingerprint else 0.0


def _position_proximity(a: Block, b: Block, sheet_diag: float) -> float:
    """
    Normalized closeness of two blocks' top-left corners.

    Manhattan distance divided by the sheet's diagonal length, then
    inverted: 1.0 means same position, 0.0 means opposite corners.
    """
    if sheet_diag <= 0:
        return 1.0 if a.top_left == b.top_left else 0.0
    dr = abs(a.top_left[0] - b.top_left[0])
    dc = abs(a.top_left[1] - b.top_left[1])
    distance = dr + dc
    return max(0.0, 1.0 - distance / sheet_diag)


def _neighbor_context_overlap(a: Block, b: Block) -> float:
    """Jaccard similarity of the two blocks' neighbor_context label sets."""
    sa = set(a.neighbor_context)
    sb = set(b.neighbor_context)
    if not sa and not sb:
        return 1.0      # both have no neighbors → trivially "same context"
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / len(sa | sb)


def _shape_similarity(a: Block, b: Block) -> float:
    """
    1.0 if shapes identical, decaying linearly with relative size delta.
    """
    ra, ca = a.shape
    rb, cb = b.shape
    row_sim = 1.0 - abs(ra - rb) / max(ra, rb, 1)
    col_sim = 1.0 - abs(ca - cb) / max(ca, cb, 1)
    return (row_sim + col_sim) / 2.0


def _label_set_jaccard(a: Block, b: Block) -> float:
    """Jaccard similarity over the full label sets of two blocks."""
    sa = set(a.labels)
    sb = set(b.labels)
    if not sa and not sb:
        return 1.0
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / len(sa | sb)


def _composite_score(
    a: Block,
    b: Block,
    sheet_diag: float,
) -> tuple[float, dict[str, float]]:
    """
    Compute the weighted composite match score.

    Returns the total score and the per-component breakdown for
    debugging / reporting.
    """
    components = {
        "label_similarity": _label_similarity(a, b),
        "fingerprint_match": _fingerprint_match(a, b),
        "position_proximity": _position_proximity(a, b, sheet_diag),
        "neighbor_context": _neighbor_context_overlap(a, b),
        "shape_similarity": _shape_similarity(a, b),
    }
    total = sum(components[k] * config.MATCH_WEIGHTS[k] for k in components)
    return total, components


def _sheet_diagonal(blocks: list[Block]) -> float:
    """Estimate sheet diagonal from block extents (used for normalizing distance)."""
    if not blocks:
        return 1.0
    max_row = max(b.bottom_right[0] for b in blocks)
    max_col = max(b.bottom_right[1] for b in blocks)
    return float(max_row + max_col) or 1.0


# ── Pass 1 — exact fingerprint ────────────────────────────────────────


def _pass1_fingerprint(
    old_blocks: list[Block],
    new_blocks: list[Block],
    sheet_diag: float,
) -> tuple[list[MatchPair], set[str], set[str]]:
    """
    Match blocks whose fingerprints are identical.

    Returns (pairs, matched_old_ids, matched_new_ids).
    """
    pairs: list[MatchPair] = []
    matched_old: set[str] = set()
    matched_new: set[str] = set()

    # Group new blocks by fingerprint for fast lookup
    new_by_fp: dict[str, list[Block]] = {}
    for nb in new_blocks:
        new_by_fp.setdefault(nb.fingerprint, []).append(nb)

    for ob in old_blocks:
        candidates = new_by_fp.get(ob.fingerprint, [])
        # Filter out already-matched candidates
        candidates = [c for c in candidates if c.block_id not in matched_new]
        if not candidates:
            continue

        if len(candidates) == 1:
            chosen = candidates[0]
        else:
            # Multiple new blocks share this fingerprint — pick the closest
            chosen = min(
                candidates,
                key=lambda c: abs(c.top_left[0] - ob.top_left[0])
                              + abs(c.top_left[1] - ob.top_left[1]),
            )

        confidence = 0.95
        _, breakdown = _composite_score(ob, chosen, sheet_diag)
        pairs.append(MatchPair(
            old_block=ob,
            new_block=chosen,
            confidence=confidence,
            pass_used=1,
            needs_review=False,
            score_breakdown=breakdown,
        ))
        matched_old.add(ob.block_id)
        matched_new.add(chosen.block_id)

    return pairs, matched_old, matched_new


# ── Pass 2 — hybrid composite score ───────────────────────────────────


def _pass2_composite(
    old_blocks: list[Block],
    new_blocks: list[Block],
    matched_old: set[str],
    matched_new: set[str],
    sheet_diag: float,
) -> list[MatchPair]:
    """
    Greedy hybrid matching for blocks not paired in pass 1.

    Builds the full score matrix between unmatched old and new blocks,
    then assigns pairs in descending score order. Pairs above the auto-
    accept threshold are taken; pairs in the review zone are taken but
    flagged; pairs below the reject threshold are dropped.
    """
    remaining_old = [b for b in old_blocks if b.block_id not in matched_old]
    remaining_new = [b for b in new_blocks if b.block_id not in matched_new]
    if not remaining_old or not remaining_new:
        return []

    # Build candidate list with scores
    candidates: list[tuple[float, dict[str, float], Block, Block]] = []
    for ob in remaining_old:
        for nb in remaining_new:
            score, breakdown = _composite_score(ob, nb, sheet_diag)
            if score >= config.REVIEW_ZONE_THRESHOLD * 0.6:
                # Pre-filter very-low scores so the greedy loop is cheap
                candidates.append((score, breakdown, ob, nb))

    candidates.sort(key=lambda t: t[0], reverse=True)

    pairs: list[MatchPair] = []
    used_old: set[str] = set()
    used_new: set[str] = set()

    for score, breakdown, ob, nb in candidates:
        if ob.block_id in used_old or nb.block_id in used_new:
            continue
        if score < config.REVIEW_ZONE_THRESHOLD:
            continue

        pairs.append(MatchPair(
            old_block=ob,
            new_block=nb,
            confidence=score,
            pass_used=2,
            needs_review=score < config.AUTO_ACCEPT_THRESHOLD,
            score_breakdown=breakdown,
        ))
        used_old.add(ob.block_id)
        used_new.add(nb.block_id)

    return pairs


# ── Pass 3 — label set Jaccard fallback ───────────────────────────────


def _pass3_label_jaccard(
    old_blocks: list[Block],
    new_blocks: list[Block],
    already_matched_old: set[str],
    already_matched_new: set[str],
    sheet_diag: float,
) -> list[MatchPair]:
    """
    Catch blocks that share most internal labels but had primary label
    drift big enough to fail the composite score in pass 2.
    """
    remaining_old = [b for b in old_blocks if b.block_id not in already_matched_old]
    remaining_new = [b for b in new_blocks if b.block_id not in already_matched_new]
    if not remaining_old or not remaining_new:
        return []

    # Same greedy approach but using label-set Jaccard combined with
    # position proximity as a tiebreaker.
    candidates: list[tuple[float, Block, Block]] = []
    for ob in remaining_old:
        for nb in remaining_new:
            jacc = _label_set_jaccard(ob, nb)
            if jacc < 0.5:
                continue
            pos = _position_proximity(ob, nb, sheet_diag)
            # Combined score: Jaccard dominant, position as tiebreaker
            combined = 0.7 * jacc + 0.3 * pos
            candidates.append((combined, ob, nb))

    candidates.sort(key=lambda t: t[0], reverse=True)

    pairs: list[MatchPair] = []
    used_old: set[str] = set()
    used_new: set[str] = set()

    for score, ob, nb in candidates:
        if ob.block_id in used_old or nb.block_id in used_new:
            continue
        # Cap pass-3 confidence so reviewers know these are weaker matches
        confidence = min(score, 0.70)
        _, breakdown = _composite_score(ob, nb, sheet_diag)
        pairs.append(MatchPair(
            old_block=ob,
            new_block=nb,
            confidence=confidence,
            pass_used=3,
            needs_review=True,
            score_breakdown=breakdown,
        ))
        used_old.add(ob.block_id)
        used_new.add(nb.block_id)

    return pairs


# ── Public API ────────────────────────────────────────────────────────


def match_blocks(
    old_blocks: list[Block],
    new_blocks: list[Block],
) -> MatchResult:
    """
    Match blocks across two snapshots of the same sheet.

    Args:
        old_blocks: Blocks from the older snapshot.
        new_blocks: Blocks from the newer snapshot.

    Returns:
        MatchResult with paired blocks and unmatched leftovers.
    """
    if not old_blocks and not new_blocks:
        return MatchResult()

    sheet_diag = max(_sheet_diagonal(old_blocks), _sheet_diagonal(new_blocks))

    pairs1, matched_old, matched_new = _pass1_fingerprint(
        old_blocks, new_blocks, sheet_diag
    )

    pairs2 = _pass2_composite(
        old_blocks, new_blocks, matched_old, matched_new, sheet_diag
    )
    for p in pairs2:
        matched_old.add(p.old_block.block_id)
        matched_new.add(p.new_block.block_id)

    pairs3 = _pass3_label_jaccard(
        old_blocks, new_blocks, matched_old, matched_new, sheet_diag
    )
    for p in pairs3:
        matched_old.add(p.old_block.block_id)
        matched_new.add(p.new_block.block_id)

    all_pairs = pairs1 + pairs2 + pairs3
    unmatched_old = [b for b in old_blocks if b.block_id not in matched_old]
    unmatched_new = [b for b in new_blocks if b.block_id not in matched_new]

    logger.debug(
        "Match result: pass1=%d pass2=%d pass3=%d unmatched_old=%d unmatched_new=%d",
        len(pairs1), len(pairs2), len(pairs3),
        len(unmatched_old), len(unmatched_new),
    )

    return MatchResult(
        pairs=all_pairs,
        unmatched_old=unmatched_old,
        unmatched_new=unmatched_new,
    )


# ── Sheet matching (for sheet renames) ────────────────────────────────


def match_sheet_names(
    old_names: list[str],
    new_names: list[str],
    old_fingerprint_sets: dict[str, set[str]],
    new_fingerprint_sets: dict[str, set[str]],
) -> tuple[dict[str, str], list[str], list[str]]:
    """
    Match sheets between two workbooks, detecting renames.

    Args:
        old_names: Sheet names from old workbook (in order).
        new_names: Sheet names from new workbook (in order).
        old_fingerprint_sets: {sheet_name: set of block fingerprints}
        new_fingerprint_sets: {sheet_name: set of block fingerprints}

    Returns:
        (matches, unmatched_old, unmatched_new) where matches is a dict
        mapping old_name → new_name (includes exact and renamed pairs).
    """
    matches: dict[str, str] = {}
    used_new: set[str] = set()

    # Pass A — exact name matches
    new_set = set(new_names)
    for name in old_names:
        if name in new_set:
            matches[name] = name
            used_new.add(name)

    # Pass B — fuzzy name + content overlap, combined score
    #
    # The two signals are blended into a single score (50/50) and
    # compared against one threshold. This catches both cosmetic
    # renames (high name fuzz, modest overlap) and abbreviation
    # renames like "Calcs" → "Calculations" (low name fuzz, high
    # overlap), which the previous two-gate logic missed.
    remaining_old = [n for n in old_names if n not in matches]
    remaining_new = [n for n in new_names if n not in used_new]

    # Threshold for the combined score — derived from the two underlying
    # thresholds so changing either still has the expected effect.
    combined_threshold = (
            config.SHEET_NAME_FUZZY_THRESHOLD / 100.0 * 0.5
            + config.SHEET_CONTENT_OVERLAP_THRESHOLD * 0.5
    )

    for old_name in remaining_old:
        old_fps = old_fingerprint_sets.get(old_name, set())
        best_candidate: str | None = None
        best_score: float = 0.0

        for new_name in remaining_new:
            if new_name in used_new:
                continue
            name_ratio = fuzz.ratio(old_name, new_name) / 100.0
            new_fps = new_fingerprint_sets.get(new_name, set())
            if old_fps or new_fps:
                overlap = len(old_fps & new_fps) / max(len(old_fps | new_fps), 1)
            else:
                overlap = 0.0

            combined = 0.5 * name_ratio + 0.5 * overlap
            if combined < combined_threshold:
                continue
            if combined > best_score:
                best_score = combined
                best_candidate = new_name

        if best_candidate:
            matches[old_name] = best_candidate
            used_new.add(best_candidate)

    unmatched_old = [n for n in old_names if n not in matches]
    unmatched_new = [n for n in new_names if n not in used_new]
    return matches, unmatched_old, unmatched_new
