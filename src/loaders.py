"""
src/loaders.py
--------------
All data ingestion logic. Notebooks call load_claims_file() — no pd.read_*
calls outside this module.
"""
import logging
from pathlib import Path
from typing import List, Optional

import pandas as pd

log = logging.getLogger(__name__)


# ── Column normalisation helpers ───────────────────────────────────────────

_COLUMN_MAP = {
    "Warehouse":      ["warehouse", "wh", "facility", "site", "location", "bucket name", "bucket"],
    "Start Date":     ["start date", "open date", "opened date", "date opened",
                       "created date", "create date", "claim open date"],
    "Completed Date": ["completed date", "close date", "closed date",
                       "completion date", "resolved date"],
    "Task Name":      ["task name", "task", "claim", "description", "title", "name"],
    "Labels":         ["labels", "label", "tags", "tag", "category", "categories"],
    "Due Date":       ["due date", "due", "target date", "task due", "deadline"],
}


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Return the first column in df whose name matches any candidate (case-insensitive)."""
    lower_map = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardise column names and dtypes across input files.

    Canonical schema after normalization:
        Warehouse, Start Date, Completed Date, Task Name, Labels, Due Date
    """
    df = df.copy()

    rename = {}
    for canonical, candidates in _COLUMN_MAP.items():
        found = find_column(df, candidates)
        if found and found != canonical:
            rename[found] = canonical
    df.rename(columns=rename, inplace=True)

    for col in ("Start Date", "Completed Date", "Due Date"):
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in df.select_dtypes(include="object").columns:
        df[col] = df[col].str.strip()

    return df


# ── Public loader ──────────────────────────────────────────────────────────

def load_claims_file(filepath: Path) -> pd.DataFrame:
    """
    Read a Claims Planner Excel file and return a normalised DataFrame.

    Args:
        filepath: Absolute path to the .xlsx file.

    Returns:
        DataFrame with canonical column names and coerced date types.

    Raises:
        FileNotFoundError: if filepath does not exist.
    """
    if not filepath.exists():
        raise FileNotFoundError(f"Claims file not found: {filepath}")

    log.info("Loading: %s", filepath.name)
    xl = pd.ExcelFile(filepath, engine="openpyxl")
    sheet = "Consolidated Data" if "Consolidated Data" in xl.sheet_names else xl.sheet_names[0]
    df = xl.parse(sheet)
    df = normalize_columns(df)
    log.info("  → %d rows, %d columns", len(df), len(df.columns))
    return df
