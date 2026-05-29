"""
Column mapping: target schema definition and DataFrame transformation.
"""
import pandas as pd

TARGET_COLUMNS = [
    "קוד",
    "תיאור",
    "תמונה",
    "כמות",
    "מחיר יחידה כולל מע״מ",
    "אחוז הנחה",
    "סה״כ",
]

# Columns that should have data in most rows; violations produce a warning
# (not a hard block — the user can still export).
REQUIRED_COLUMNS = ["תיאור", "כמות", "מחיר יחידה כולל מע״מ"]

NO_MAP_OPTION = "לא למפות"

# A row is considered "empty" when its value is one of these after stripping.
_EMPTY_VALUES = {"", "none", "nan", "null", "-"}
_MOSTLY_EMPTY_THRESHOLD = 0.8   # warn when ≥80 % of rows are empty


def _is_empty(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    return str(value).strip().lower() in _EMPTY_VALUES


def validate_mapping(source_df: pd.DataFrame, mapping: dict) -> list:
    """
    Return a list of Hebrew warning strings (empty list = no warnings).
    Covers:
      - required columns that are not mapped
      - required columns that are mapped but mostly empty in the source data
    Warnings are informational only; the caller decides whether to block.
    """
    warnings = []
    n = len(source_df)

    for col in REQUIRED_COLUMNS:
        source_col = mapping.get(col, NO_MAP_OPTION)

        if not source_col or source_col == NO_MAP_OPTION:
            warnings.append(f'עמודת "{col}" אינה ממופה.')
            continue

        if source_col not in source_df.columns:
            # shouldn't happen in normal use but guard defensively
            warnings.append(f'עמודת המקור "{source_col}" לא נמצאה בקובץ.')
            continue

        empty_count = source_df[source_col].apply(_is_empty).sum()
        if n > 0 and empty_count / n >= _MOSTLY_EMPTY_THRESHOLD:
            warnings.append(
                f'עמודת "{col}" ריקה ב-{empty_count} מתוך {n} שורות.'
            )

    return warnings


def apply_mapping(source_df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """
    Build a new DataFrame with exactly TARGET_COLUMNS in order.

    - Unmapped columns produce an empty column (no crash).
    - NaN / None values in source data are converted to empty strings.
    """
    n = len(source_df)
    result = {}

    for target_col in TARGET_COLUMNS:
        source_col = mapping.get(target_col, NO_MAP_OPTION)
        if source_col and source_col != NO_MAP_OPTION and source_col in source_df.columns:
            series = (
                source_df[source_col]
                .reset_index(drop=True)
                .apply(lambda v: "" if _is_empty(v) else str(v))
            )
            result[target_col] = series
        else:
            result[target_col] = pd.Series([""] * n, dtype=str)

    return pd.DataFrame(result)
