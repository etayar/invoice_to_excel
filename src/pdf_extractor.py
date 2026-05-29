"""
Word-position-based PDF table extractor.

Primary strategy for Hebrew RTL invoices (with or without table borders).

Architecture
------------
1. collect_rows   – extract_words() per page; skip chrome (header/footer).
2. detect_columns – frequency-based "anchor" detection across ALL rows.
   Columns whose x-centre appears in ≥ ANCHOR_FREQ of rows are "anchored"
   (numbers, barcodes). Wide gaps that contain scattered words get a
   synthetic "description" column inserted between the anchors.
3. assign words   – bucket each word into its column by x-centre.
   Hebrew characters are reversed (this PDF stores them in visual order).
4. split header   – first row with ≥ 2 numeric cells = start of data;
   everything before it is merged into a composite header row.

Debug mode
----------
Call extract_table(pdf, debug_dir="debug_output") to write:
  • debug_words.csv      – every extracted word with coordinates
  • debug_col_breaks.txt – detected column break positions
  • debug_rows.csv       – every reconstructed row (cells)
"""

from __future__ import annotations

import csv
import os
from collections import Counter, defaultdict

import pandas as pd
import pdfplumber

# ---------------------------------------------------------------------------
# Tuning constants
# ---------------------------------------------------------------------------
_META_TOP      = 280   # pt from page top — company/invoice metadata area above
_FOOTER_TOP    = 780   # pt from page top — page-number / timestamp area below
_Y_TOLERANCE   = 4     # words within this many pt share a row
_COL_BUCKET    = 8     # pt width for x-centre quantisation
_ANCHOR_FREQ   = 0.60  # fraction of rows an x-bin must appear in to be an anchor
# NOTE: true numeric/barcode columns appear in 70–100 % of rows.
#       Repeated brand-name / layout-coincidence words peak at 55 %, so 0.60
#       cleanly separates signal from noise across both large and small tables.
_ANCHOR_SPREAD = 5.0   # max pt spread (max-min x) within an anchor bucket
# NOTE: numbers in a true column always land at the same x (spread≈0).
#       Description words at "consistent" positions (due to regular layout)
#       still drift ≥6 pt across rows — this filter removes them.
_ANCHOR_MERGE  = 12    # pt — merge two anchors closer than this
_WIDE_GAP      = 40    # pt — gaps wider than this may hide a description column
_MIN_COLS      = 2     # reject result if fewer columns than this
_MIN_GAP       = 8     # minimum pt gap to register a column break (first-row fallback)


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def extract_table(pdf: pdfplumber.PDF, debug_dir: str | None = None):
    """
    Extract the main data table from *pdf*.
    Returns a DataFrame or None if extraction fails.
    Writes debug files to *debug_dir* when provided.

    Tries two metadata-skip thresholds in order:
      280 pt — skips the invoice header block present in real-world invoices
       50 pt — falls back for simple PDFs where the table starts near the top
    The first attempt that produces a usable result is returned.
    """
    page_width = pdf.pages[0].width if pdf.pages else 595.0

    for meta_top in (_META_TOP, 50):
        all_rows = _collect_rows(pdf, meta_top=meta_top)
        if not all_rows:
            continue

        # Detect whether Hebrew characters are stored in visual (reversed) order.
        # Older Israeli accounting software encodes Hebrew visually; modern tools
        # and reportlab use logical Unicode order.
        needs_rev = _needs_hebrew_reversal(all_rows)

        col_breaks = _detect_columns(all_rows, page_width)
        if len(col_breaks) < _MIN_COLS + 1:
            continue

        table_rows = _build_table_rows(all_rows, col_breaks, reverse_hebrew=needs_rev)
        if len(table_rows) < 2:
            continue

        header, data = _split_header_data(table_rows)
        if not data:
            continue

        df = _build_df(header, data)
        if df is None or df.empty or len(df.columns) < _MIN_COLS:
            continue

        # Require either named columns or a substantial number of data rows.
        # A fragment of 2 rows with all-generic column names is not good enough
        # — keep trying the next meta_top in case a better result is available.
        has_named = any(not c.startswith("עמודה_") for c in df.columns)
        if not has_named and len(df) < 5:
            continue

        if debug_dir:
            _debug_write_words(all_rows, debug_dir)
            _debug_write_breaks(col_breaks, debug_dir)
            _debug_write_rows(table_rows, debug_dir)
        return df

    return None


# ---------------------------------------------------------------------------
# Phase 1 — collect rows
# ---------------------------------------------------------------------------

def _collect_rows(
    pdf: pdfplumber.PDF,
    meta_top: float = _META_TOP,
) -> list[list[dict]]:
    all_rows: list[list[dict]] = []

    for page in pdf.pages:
        words = page.extract_words(
            x_tolerance=3,
            y_tolerance=3,
            keep_blank_chars=False,
            use_text_flow=False,
        )
        if not words:
            continue

        for row_words in _group_by_y(words):
            if not row_words:
                continue
            avg_y = sum(w["top"] for w in row_words) / len(row_words)
            # Skip chrome: metadata header (above meta_top) and footer
            if meta_top <= avg_y <= _FOOTER_TOP:
                all_rows.append(row_words)

    return all_rows


def _group_by_y(words: list[dict], tol: int = _Y_TOLERANCE) -> list[list[dict]]:
    if not words:
        return []
    sw = sorted(words, key=lambda w: (w["top"], w["x0"]))
    rows: list[list[dict]] = []
    cur, cur_top = [sw[0]], sw[0]["top"]

    for w in sw[1:]:
        if abs(w["top"] - cur_top) <= tol:
            cur.append(w)
        else:
            rows.append(sorted(cur, key=lambda w: w["x0"]))
            cur, cur_top = [w], w["top"]

    rows.append(sorted(cur, key=lambda w: w["x0"]))
    return rows


# ---------------------------------------------------------------------------
# Phase 2 — detect columns
# ---------------------------------------------------------------------------

def _detect_columns(rows: list[list[dict]], page_width: float) -> list[float]:
    """
    Frequency-based anchor detection.

    Words in numeric/barcode columns appear at CONSISTENT x-positions across
    many rows.  Description words scatter — their x-bins have low frequency.
    We treat high-frequency x-bins as column anchors, find the gaps between
    them, and insert a synthetic "description" column in wide gaps that
    actually contain non-anchor words.
    """
    n = len(rows)

    # Track count AND actual x-centres per bucket (to compute spread).
    bucket_centers: dict[int, list[float]] = {}

    for row in rows:
        seen: dict[int, float] = {}
        for w in row:
            c = (w["x0"] + w["x1"]) / 2
            b = round(c / _COL_BUCKET) * _COL_BUCKET
            if b not in seen:
                seen[b] = c
        for b, c in seen.items():
            bucket_centers.setdefault(b, []).append(c)

    min_count = max(2, n * _ANCHOR_FREQ)
    raw_anchors = sorted(
        b for b, xs in bucket_centers.items()
        if len(xs) >= min_count
        and (max(xs) - min(xs)) <= _ANCHOR_SPREAD   # tight spread → true column
    )

    # Fall back to first-row gap analysis if no anchors found
    if not raw_anchors:
        return _first_row_breaks(rows[0], page_width)

    # Merge nearby anchors into single representative positions
    anchors: list[float] = [float(raw_anchors[0])]
    for x in raw_anchors[1:]:
        if x - anchors[-1] <= _ANCHOR_MERGE:
            anchors[-1] = (anchors[-1] + x) / 2
        else:
            anchors.append(float(x))

    if len(anchors) < _MIN_COLS:
        return _first_row_breaks(rows[0], page_width)

    # Compute actual word extents (min x0, max x1) per anchor cluster
    extents: dict[float, list[float]] = {}
    for row in rows:
        for w in row:
            c = (w["x0"] + w["x1"]) / 2
            for a in anchors:
                if abs(c - a) <= _ANCHOR_MERGE:
                    if a not in extents:
                        extents[a] = [w["x0"], w["x1"]]
                    else:
                        extents[a][0] = min(extents[a][0], w["x0"])
                        extents[a][1] = max(extents[a][1], w["x1"])
                    break

    # Build break list
    breaks: list[float] = [0.0]

    for i in range(len(anchors) - 1):
        a1, a2 = anchors[i], anchors[i + 1]
        right1 = extents.get(a1, [a1, a1])[1]
        left2  = extents.get(a2, [a2, a2])[0]
        gap    = left2 - right1

        if gap > _WIDE_GAP:
            # Count non-anchor words in this gap to decide whether to insert
            # a description column.
            gap_count = sum(
                1
                for row in rows
                for w in row
                if right1 + 5 < (w["x0"] + w["x1"]) / 2 < left2 - 5
                and not any(abs((w["x0"] + w["x1"]) / 2 - a) <= _ANCHOR_MERGE
                            for a in anchors)
            )
            if gap_count >= n * 0.20:
                margin = min(gap * 0.12, 10.0)
                breaks.append(right1 + margin)
                breaks.append(left2  - margin)
            else:
                breaks.append((right1 + left2) / 2)
        else:
            breaks.append((right1 + left2) / 2)

    breaks.append(float(page_width))
    return sorted(set(round(b, 1) for b in breaks if 0 <= b <= page_width))


def _first_row_breaks(row: list[dict], page_width: float) -> list[float]:
    """Fallback: detect breaks from a single row's word gaps."""
    sw = sorted(row, key=lambda w: w["x0"])
    breaks = [0.0]
    for i in range(len(sw) - 1):
        gap_start = sw[i]["x1"]
        gap_end   = sw[i + 1]["x0"]
        if gap_end - gap_start >= _MIN_GAP:
            breaks.append((gap_start + gap_end) / 2)
    breaks.append(float(page_width))
    return breaks


# ---------------------------------------------------------------------------
# Phase 2b — detect Hebrew encoding style
# ---------------------------------------------------------------------------

# Common Hebrew words that appear in invoices — used for encoding detection.
_HEBREW_INVOICE_WORDS: frozenset[str] = frozenset({
    "קוד", "מחיר", "כמות", "תיאור", "הנחה", "סה\"כ", "יחידות", "מוצר",
    "שם", "ברקוד", "סכום", "כולל", "אחוז", "יחידה", "מע\"מ", "חשבונית",
    "ריחמ", "תומכ", "רואית", "דוק",  # reversed forms — not real words
})


def _needs_hebrew_reversal(rows: list[list[dict]]) -> bool:
    """
    Return True if the PDF stores Hebrew in visual (reversed) byte order.
    Samples the first 30 rows: if more words match known Hebrew invoice terms
    when reversed than when read as-is, reversal is needed.
    """
    normal_hits   = 0
    reversed_hits = 0

    for row in rows[:30]:
        for w in row:
            text = w["text"]
            if not _is_rtl(text) or len(text) < 2:
                continue
            # Check both orientations against the known-word set
            if text in _HEBREW_INVOICE_WORDS:
                normal_hits += 1
            if text[::-1] in _HEBREW_INVOICE_WORDS:
                reversed_hits += 1

    # If reversed forms are more recognisable, reversal is required.
    return reversed_hits > normal_hits


# ---------------------------------------------------------------------------
# Phase 3 — assign words to cells
# ---------------------------------------------------------------------------

def _build_table_rows(
    all_rows: list[list[dict]],
    col_breaks: list[float],
    reverse_hebrew: bool = False,
) -> list[list[str]]:
    result: list[list[str]] = []
    for row_words in all_rows:
        cells = _words_to_cells(row_words, col_breaks, reverse_hebrew)
        if any(c.strip() for c in cells):
            result.append(cells)
    return result


def _words_to_cells(
    row_words: list[dict],
    col_breaks: list[float],
    reverse_hebrew: bool = False,
) -> list[str]:
    n_cols = len(col_breaks) - 1
    buckets: list[list[dict]] = [[] for _ in range(n_cols)]

    for w in row_words:
        c   = (w["x0"] + w["x1"]) / 2
        idx = n_cols - 1
        for i in range(n_cols):
            if col_breaks[i] <= c < col_breaks[i + 1]:
                idx = i
                break
        buckets[idx].append(w)

    return [
        _cell_text(sorted(b, key=lambda w: w["x0"]), reverse_hebrew)
        for b in buckets
    ]


def _is_rtl(text: str) -> bool:
    """True when *text* contains enough Hebrew / Arabic characters."""
    rtl = sum(1 for c in text if "א" <= c <= "߿")
    return rtl >= 2 or (len(text) > 0 and rtl / len(text) > 0.3)


def _fix_word(text: str, reverse: bool) -> str:
    """Reverse Hebrew word characters when the PDF uses visual byte order."""
    return text[::-1] if (reverse and _is_rtl(text)) else text


def _cell_text(words: list[dict], reverse_hebrew: bool = False) -> str:
    """Reconstruct a readable cell string from a bucket of words."""
    if not words:
        return ""
    texts = [_fix_word(w["text"], reverse_hebrew) for w in words]
    # For RTL-dominant cells with reversal active, flip word order so the
    # phrase reads correctly (e.g. ['יחידה','מחיר'] → 'מחיר יחידה').
    if reverse_hebrew and sum(1 for t in texts if _is_rtl(t)) > len(texts) / 2:
        texts = texts[::-1]
    return " ".join(texts)


# ---------------------------------------------------------------------------
# Phase 4 — split header / data
# ---------------------------------------------------------------------------

def _count_numeric(row: list[str]) -> int:
    return sum(
        1 for c in row
        if c.strip()
        and c.strip().replace(".", "").replace(",", "").replace("-", "").isdigit()
    )


def _split_header_data(
    table_rows: list[list[str]],
) -> tuple[list[str], list[list[str]]]:
    """
    Find where data begins (first row with ≥ 2 numeric cells).
    Everything before that is merged into a composite header.
    """
    n_cols     = len(table_rows[0]) if table_rows else 1
    data_start = 0

    for i, row in enumerate(table_rows):
        if _count_numeric(row) >= 2:
            data_start = i
            break

    header_rows = table_rows[:data_start]
    data_rows   = table_rows[data_start:]

    if not header_rows:
        header = [f"עמודה_{i + 1}" for i in range(n_cols)]
    else:
        # Merge: per column, concatenate all non-empty values from all header rows
        header = [""] * n_cols
        for hrow in header_rows:
            for j, cell in enumerate(hrow):
                if j < n_cols and cell.strip():
                    header[j] = (header[j] + " " + cell).strip() if header[j] else cell
        header = [c or f"עמודה_{i + 1}" for i, c in enumerate(header)]

    # Remove rows that exactly repeat the raw first header row (page repeats)
    if data_start > 0:
        raw_header = table_rows[data_start - 1]
        data_rows = [r for r in data_rows if r != raw_header]

    return header, data_rows


# ---------------------------------------------------------------------------
# Shared builder
# ---------------------------------------------------------------------------

def _build_df(header: list[str], data_rows: list[list[str]]) -> pd.DataFrame:
    n = len(header)
    normalised = []
    for row in data_rows:
        if len(row) < n:
            row = row + [""] * (n - len(row))
        elif len(row) > n:
            row = row[:n]
        normalised.append(row)
    return pd.DataFrame(normalised, columns=header)


# ---------------------------------------------------------------------------
# Debug helpers
# ---------------------------------------------------------------------------

def _ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def _debug_write_words(all_rows: list[list[dict]], debug_dir: str) -> None:
    _ensure_dir(debug_dir)
    path = os.path.join(debug_dir, "debug_words.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["row_idx", "x0", "x1", "top", "bottom", "text", "fixed"])
        for ri, row_words in enumerate(all_rows):
            for word in row_words:
                w.writerow([
                    ri,
                    round(word["x0"], 1), round(word["x1"], 1),
                    round(word["top"], 1), round(word["bottom"], 1),
                    word["text"],
                    _fix_word(word["text"]),
                ])
    print(f"[debug] words → {path}")


def _debug_write_breaks(col_breaks: list[float], debug_dir: str) -> None:
    _ensure_dir(debug_dir)
    path = os.path.join(debug_dir, "debug_col_breaks.txt")
    with open(path, "w", encoding="utf-8") as f:
        f.write(f"Column count : {len(col_breaks) - 1}\n")
        f.write(f"Break points : {col_breaks}\n\n")
        for i in range(len(col_breaks) - 1):
            f.write(f"  col {i + 1}: x = {col_breaks[i]:.1f} → {col_breaks[i+1]:.1f}\n")
    print(f"[debug] col_breaks → {path}")


def _debug_write_rows(table_rows: list[list[str]], debug_dir: str) -> None:
    _ensure_dir(debug_dir)
    path = os.path.join(debug_dir, "debug_rows.csv")
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        if not table_rows:
            return
        w = csv.writer(f)
        w.writerow([f"col_{i+1}" for i in range(len(table_rows[0]))])
        for row in table_rows:
            w.writerow(row)
    print(f"[debug] rows → {path}")
