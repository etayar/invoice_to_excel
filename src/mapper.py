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

NO_MAP_OPTION = "לא למפות"


def apply_mapping(source_df: pd.DataFrame, mapping: dict) -> pd.DataFrame:
    """
    Build a new DataFrame with exactly TARGET_COLUMNS in order.

    mapping: {target_col_name: source_col_name_or_NO_MAP_OPTION}
    Unmapped or missing columns are filled with empty strings.
    """
    n = len(source_df)
    result = {}

    for target_col in TARGET_COLUMNS:
        source_col = mapping.get(target_col, NO_MAP_OPTION)
        if source_col and source_col != NO_MAP_OPTION and source_col in source_df.columns:
            result[target_col] = source_df[source_col].reset_index(drop=True)
        else:
            result[target_col] = pd.Series([""] * n, dtype=str)

    return pd.DataFrame(result)
