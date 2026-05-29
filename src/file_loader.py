"""
File loading for PDF, Excel (.xlsx/.xls), and CSV inputs.
Designed so OCR support can be added later by replacing/extending _load_pdf.
"""
import os
import pandas as pd
import pdfplumber

from src.pdf_extractor import extract_table


class FileLoadError(Exception):
    """User-facing error raised when a file cannot be loaded."""


def load_file(path: str, sheet_name=None, pdf_debug_dir: str | None = None):
    """
    Load a supported file and return (DataFrame, sheet_names).

    sheet_names is a list[str] for Excel files with multiple sheets, else None.
    sheet_name selects which Excel sheet to load; defaults to the first.
    pdf_debug_dir: when set, the PDF extractor writes debug files there.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return _load_pdf(path, debug_dir=pdf_debug_dir), None
    elif ext in (".xlsx", ".xls"):
        return _load_excel(path, sheet_name)
    elif ext == ".csv":
        return _load_csv(path), None
    else:
        raise FileLoadError(f"סוג קובץ לא נתמך: {ext}\nנתמכים: PDF, Excel (.xlsx/.xls), CSV")


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
#
# Primary strategy: word-position-based extraction via pdf_extractor.py
#   – groups extract_words() output by y-coordinate
#   – detects columns from x-frequency across ALL rows (not just the first)
#   – handles Hebrew RTL (characters stored in visual order → reversed)
#   – multi-page: collects all pages before detecting column structure
#
# OCR hook point: replace or extend _load_pdf for scanned PDFs.
# ---------------------------------------------------------------------------

def _load_pdf(path: str, debug_dir: str | None = None) -> pd.DataFrame:
    """
    Extract table from PDF using word-position analysis.
    Raises FileLoadError with a user-friendly message if extraction fails.
    """
    with pdfplumber.open(path) as pdf:
        df = extract_table(pdf, debug_dir=debug_dir)

    if df is None or df.empty:
        raise FileLoadError(
            "לא ניתן לזהות טבלה בקובץ ה-PDF.\n\n"
            "ייתכן שמדובר ב-PDF סרוק הדורש זיהוי תווים (OCR),\n"
            "או שהקובץ אינו מכיל טבלה בפורמט הנתמך.\n"
            "נסה לייצא את הקובץ לאקסל ולהעלות אותו במקום."
        )

    if len(df.columns) < 2:
        raise FileLoadError(
            "הקובץ מכיל עמודה אחת בלבד — לא ניתן לבצע מיפוי.\n"
            "ודא שה-PDF מכיל טבלה עם מספר עמודות."
        )

    return df


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
