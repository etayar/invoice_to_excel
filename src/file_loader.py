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
# PDF
# ---------------------------------------------------------------------------

def _load_pdf(path: str) -> pd.DataFrame:
    """
    Extract tables from every page using pdfplumber and merge into one DataFrame.
    If no tables are found, raises a friendly FileLoadError (OCR hook point).
    """
    header = None
    data_rows = []

    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            page_tables = page.extract_tables()
            for table in page_tables:
                if not table:
                    continue
                for row_idx, raw_row in enumerate(table):
                    row = [str(c).strip() if c is not None else "" for c in raw_row]
                    if header is None:
                        # First non-empty row of first table becomes the header.
                        header = [c if c else f"עמודה_{j + 1}" for j, c in enumerate(row)]
                    else:
                        # Skip rows that are an exact repeat of the header
                        # (common when a table spans multiple pages).
                        if row == header:
                            continue
                        data_rows.append(row)

    if header is None:
        raise FileLoadError(
            "לא נמצאו טבלאות בקובץ ה-PDF.\n\n"
            "ייתכן שמדובר ב-PDF סרוק הדורש זיהוי תווים (OCR).\n"
            "במקרה זה, אנא המר את הקובץ לאקסל לפני ההעלאה."
        )

    if not data_rows:
        raise FileLoadError(
            "הטבלאות שנמצאו ב-PDF אינן מכילות שורות נתונים.\n"
            "ודא שהקובץ מכיל טבלה עם לפחות שורת נתונים אחת."
        )

    # Normalise row widths to match the header length.
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
