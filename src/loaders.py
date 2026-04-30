# src/loaders.py — Data ingestion helpers for Claims Aging Pipeline.
# All pd.read_* calls live here. Notebooks never read data directly.
import logging
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

log = logging.getLogger(__name__)

_COLUMN_MAP = {
    "Warehouse":      ["bucket name", "warehouse", "wh", "facility", "site", "location"],
    "Start Date":     ["start date", "open date", "opened date", "created date",
                       "create date", "claim open date", "date opened"],
    "Completed Date": ["completed date", "close date", "closed date",
                       "completion date", "resolved date"],
    "Task Name":      ["task name", "task", "claim", "description", "title", "name"],
    "Labels":         ["labels", "label", "tags", "tag", "category", "categories"],
    "Due Date":       ["due date", "due", "target date", "task due", "deadline"],
}


# ── Utility helpers ────────────────────────────────────────────────────────

def safe_filename(name: str, max_len: int = 180) -> str:
    """Sanitize a string for use as a filesystem filename."""
    import re
    if name is None:
        name = "untitled"
    name = re.sub(r'[<>:"/\\|?*]', "_", str(name))
    name = re.sub(r"[\x00-\x1f]", "_", name)
    name = re.sub(r"_+", "_", name).strip(" .")
    return (name or "unnamed")[:max_len]


def find_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """Return the first column name (case-insensitive) that matches any candidate."""
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    return None


# ── Loaders ────────────────────────────────────────────────────────────────

def load_excel(path: Path, sheet_name: int | str = 0, **kwargs) -> pd.DataFrame:
    """Load an Excel file and return a raw DataFrame."""
    return pd.read_excel(path, sheet_name=sheet_name, **kwargs)


def load_claims_file(filepath: Path) -> pd.DataFrame:
    """
    Read a Claims Planner Excel file and return a normalised DataFrame.
    Raises FileNotFoundError if filepath does not exist.
    """
    if not filepath.exists():
        raise FileNotFoundError(f"Claims file not found: {filepath}")
    log.info("Loading: %s", filepath.name)
    df = pd.read_excel(filepath, engine="openpyxl")
    df, _ = normalize_columns(df)
    log.info("  → %d rows, %d columns", len(df), len(df.columns))
    return df


# ── Column normalization ───────────────────────────────────────────────────

def normalize_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, dict]:
    """
    Detect and standardize column names across Ford and Chrysler export schemas.

    Returns:
        (normalized_df, rename_map) — rename_map records what was renamed for audit.

    Raises:
        KeyError if required Warehouse or Start Date column cannot be found.
    """
    out = df.copy()
    rename_map: dict = {}

    for canonical, candidates in _COLUMN_MAP.items():
        found = find_column(out, candidates)
        if found and found != canonical:
            out.rename(columns={found: canonical}, inplace=True)
            rename_map[found] = canonical

    if "Warehouse" not in out.columns:
        raise KeyError("Warehouse column not found — checked: " +
                       ", ".join(_COLUMN_MAP["Warehouse"]))
    if "Start Date" not in out.columns:
        raise KeyError("Start/Open date column not found — checked: " +
                       ", ".join(_COLUMN_MAP["Start Date"]))

    for col in ("Completed Date", "Due Date"):
        if col not in out.columns:
            out[col] = pd.NaT
    for col in ("Task Name", "Labels"):
        if col not in out.columns:
            out[col] = ""

    out["Start Date"]     = pd.to_datetime(out["Start Date"],     errors="coerce")
    out["Completed Date"] = pd.to_datetime(out["Completed Date"], errors="coerce")
    out["Due Date"]       = pd.to_datetime(out["Due Date"],        errors="coerce")
    out["Warehouse"]      = out["Warehouse"].astype(str).str.strip()
    out["Task Name"]      = out["Task Name"].astype(str).str.strip()
    out["Labels"]         = out["Labels"].astype(str)

    return out, rename_map
