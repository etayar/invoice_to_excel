"""
File loading for PDF, Excel (.xlsx/.xls), and CSV inputs.
Designed so OCR support can be added later by replacing/extending _load_pdf.
"""
import os
import pandas as pd
import pdfplumber


class FileLoadError(Exception):
    """User-facing error raised when a file cannot be loaded."""


def load_file(path: str, sheet_name=None):
    """
    Load a supported file and return (DataFrame, sheet_names).

    sheet_names is a list[str] for Excel files with multiple sheets, else None.
    sheet_name selects which Excel sheet to load; defaults to the first.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return _load_pdf(path), None
    elif ext in (".xlsx", ".xls"):
        return _load_excel(path, sheet_name)
    elif ext == ".csv":
        return _load_csv(path), None
    else:
        raise FileLoadError(f"סוג קובץ לא נתמך: {ext}\nנתמכים: PDF, Excel (.xlsx/.xls), CSV")


# ---------------------------------------------------------------------------
# PDF — three-strategy cascade
# ---------------------------------------------------------------------------
#
# Root causes of the reported bugs:
#
# 1. Borderless PDFs (most common in Israeli invoice software):
#    pdfplumber's default extract_tables() needs line/border elements to split
#    columns. Without them it returns every row as a single merged cell.
#    Fix: fall back to word-position-based column detection (_strategy_words).
#
# 2. עמודה_N generic column names:
#    The first row of a detected table was all-empty (spacer / background row).
#    The old code accepted it as the header and filled gaps with עמודה_N.
#    Fix: skip all-empty rows while scanning for the first real header row.
#
# 3. Multi-page tables:
#    Repeated header rows on subsequent pages were not always skipped.
#    Fix: both strategies normalise repeated header rows.
#
# OCR hook point: replace or extend _load_pdf for scanned PDFs.
# ---------------------------------------------------------------------------

_MIN_COLS = 2          # fewer columns than this → extraction considered failed
_MIN_GAP  = 8          # points of horizontal whitespace that marks a column break


def _load_pdf(path: str) -> pd.DataFrame:
    """
    Try three strategies in order; return the first acceptable result.
    Raises FileLoadError if all strategies fail.
    """
    with pdfplumber.open(path) as pdf:
        # Strategy 1: border/line-based (fast; works when PDF has explicit gridlines)
        df = _strategy_border(pdf)
        if _is_usable(df):
            return df

        # Strategy 2: word-position-based (works for borderless tables)
        df = _strategy_words(pdf)
        if _is_usable(df):
            return df

    raise FileLoadError(
        "לא ניתן לזהות טבלה בקובץ ה-PDF.\n\n"
        "ייתכן שמדובר ב-PDF סרוק הדורש זיהוי תווים (OCR),\n"
        "או שהקובץ אינו מכיל טבלה בפורמט הנתמך.\n"
        "נסה לייצא את הקובץ לאקסל ולהעלות אותו במקום."
    )


def _is_usable(df) -> bool:
    """Return True when df looks like a real data table."""
    if df is None or df.empty:
        return False
    if len(df.columns) < _MIN_COLS:
        return False
    # At least 30 % of cells across the first 10 rows should be non-empty.
    sample = df.head(10)
    total  = sample.size
    filled = (sample != "").sum().sum()
    return total > 0 and (filled / total) >= 0.30


# ── Strategy 1: border-based ─────────────────────────────────────────────────

def _strategy_border(pdf) -> pd.DataFrame:
    header    = None
    data_rows = []

    for page in pdf.pages:
        for table in page.extract_tables():
            if not table:
                continue
            for raw_row in table:
                row = [str(c).strip() if c is not None else "" for c in raw_row]

                if all(c == "" for c in row):
                    continue                         # skip spacer / empty rows

                if header is None:
                    header = [c or f"עמודה_{j+1}" for j, c in enumerate(row)]
                elif row == header:
                    continue                         # skip repeated page headers
                else:
                    data_rows.append(row)

    if header is None or not data_rows:
        return None

    return _build_df(header, data_rows)


# ── Strategy 2: word-position-based ──────────────────────────────────────────

def _strategy_words(pdf) -> pd.DataFrame:
    """
    Reconstruct table columns from the x-positions of words on each page.
    Works for borderless tables where spacing alone separates columns.
    """
    col_breaks = None   # detected from the first (header) row
    header     = None
    data_rows  = []

    for page in pdf.pages:
        words = page.extract_words(
            x_tolerance=3,
            y_tolerance=3,
            keep_blank_chars=False,
            use_text_flow=False,
        )
        if not words:
            continue

        row_groups = _group_words_by_y(words, y_tolerance=4)

        for word_row in row_groups:
            if not word_row:
                continue

            if col_breaks is None:
                # Use the first non-trivial row to set column structure.
                if len(word_row) < _MIN_COLS:
                    continue
                col_breaks = _find_col_breaks(word_row, page.width)
                if len(col_breaks) < _MIN_COLS + 1:
                    col_breaks = None
                    continue

            row = _words_to_row(word_row, col_breaks)

            if all(c == "" for c in row):
                continue

            if header is None:
                header = [c or f"עמודה_{j+1}" for j, c in enumerate(row)]
            elif row == header:
                continue                             # skip repeated page headers
            else:
                data_rows.append(row)

    if header is None or not data_rows:
        return None

    return _build_df(header, data_rows)


def _group_words_by_y(words: list, y_tolerance: int = 4) -> list:
    """Cluster words into rows by their vertical (top) position."""
    sorted_words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    rows, current, current_top = [], [sorted_words[0]], sorted_words[0]["top"]

    for word in sorted_words[1:]:
        if abs(word["top"] - current_top) <= y_tolerance:
            current.append(word)
        else:
            rows.append(sorted(current, key=lambda w: w["x0"]))
            current, current_top = [word], word["top"]

    rows.append(sorted(current, key=lambda w: w["x0"]))
    return rows


def _find_col_breaks(header_words: list, page_width: float) -> list:
    """
    Identify column split positions from the gaps between header words.
    Returns a list of x-coordinates that bound each column:
    [left_edge, break1, break2, ..., right_edge]
    """
    sw = sorted(header_words, key=lambda w: w["x0"])
    breaks = [0.0]

    for i in range(len(sw) - 1):
        gap_start = sw[i]["x1"]
        gap_end   = sw[i + 1]["x0"]
        if gap_end - gap_start >= _MIN_GAP:
            breaks.append((gap_start + gap_end) / 2.0)

    breaks.append(page_width)
    return breaks


def _words_to_row(word_row: list, breaks: list) -> list:
    """Assign each word to its column bucket and join words within each bucket."""
    n_cols = len(breaks) - 1
    buckets = [[] for _ in range(n_cols)]

    for word in word_row:
        center = (word["x0"] + word["x1"]) / 2.0
        idx = n_cols - 1
        for i in range(n_cols):
            if breaks[i] <= center < breaks[i + 1]:
                idx = i
                break
        buckets[idx].append(word["text"])

    return [" ".join(ws) for ws in buckets]


# ── Shared helper ─────────────────────────────────────────────────────────────

def _build_df(header: list, data_rows: list) -> pd.DataFrame:
    """Normalise row widths and return a DataFrame."""
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
# Excel
# ---------------------------------------------------------------------------

def _load_excel(path: str, sheet_name=None):
    """Return (DataFrame, list_of_sheet_names)."""
    try:
        xl = pd.ExcelFile(path)
    except Exception as exc:
        raise FileLoadError(f"לא ניתן לפתוח את קובץ האקסל:\n{exc}") from exc

    sheet_names = xl.sheet_names

    if sheet_name is None or sheet_name not in sheet_names:
        sheet_name = sheet_names[0]

    try:
        df = xl.parse(sheet_name, dtype=str)
    except Exception as exc:
        raise FileLoadError(f"שגיאה בקריאת הגיליון '{sheet_name}':\n{exc}") from exc

    df.columns = [str(c) for c in df.columns]
    df = df.fillna("")
    return df, sheet_names


# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------

_CSV_ENCODINGS = ("utf-8-sig", "utf-8", "windows-1255", "cp1255", "latin-1")


def _load_csv(path: str) -> pd.DataFrame:
    for enc in _CSV_ENCODINGS:
        try:
            df = pd.read_csv(path, encoding=enc, dtype=str)
            df.columns = [str(c) for c in df.columns]
            df = df.fillna("")
            return df
        except UnicodeDecodeError:
            continue
        except Exception as exc:
            raise FileLoadError(f"שגיאה בקריאת קובץ ה-CSV:\n{exc}") from exc

    raise FileLoadError(
        "לא ניתן לקרוא את קובץ ה-CSV.\n"
        "נסה לשמור אותו מחדש בקידוד UTF-8."
    )
