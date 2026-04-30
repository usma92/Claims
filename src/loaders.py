# src/loaders.py — Data ingestion helpers for Claims Aging Pipeline.
# All pd.read_* calls live here. Notebooks never read data directly.
import re
from pathlib import Path
from typing import Optional, Tuple

import pandas as pd


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def safe_filename(name: str, max_len: int = 180) -> str:
    """Sanitize a string for use as a filesystem filename."""
    if name is None:
        name = "untitled"
    name = re.sub(r'[<>:"/\\|?*]', "_", str(name))
    name = re.sub(r"[\x00-\x1f]", "_", name)
    name = re.sub(r"_+", "_", name).strip(" .")
    if not name:
        name = "unnamed"
    return name[:max_len]


def find_column(df: pd.DataFrame, candidates: list[str]) -> Optional[str]:
    """Return the first column name (case-insensitive) that matches any candidate."""
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in cols:
            return cols[cand.lower()]
    return None


# ---------------------------------------------------------------------------
# Loaders
# ---------------------------------------------------------------------------

def load_excel(path: Path, sheet_name: int | str = 0, **kwargs) -> pd.DataFrame:
    """Load an Excel file and return a raw DataFrame."""
    return pd.read_excel(path, sheet_name=sheet_name, **kwargs)


# ---------------------------------------------------------------------------
# Column normalization
# ---------------------------------------------------------------------------

def normalize_columns(df: pd.DataFrame) -> Tuple[pd.DataFrame, dict]:
    """
    Detect and standardize column names across Ford and Chrysler export schemas.

    Returns:
        (normalized_df, rename_map)  — rename_map records what was renamed for audit.

    Raises:
        KeyError if required warehouse or start-date column cannot be found.
    """
    wh_col     = find_column(df, ["Bucket Name", "Warehouse", "WH", "Facility"])
    sd_col     = find_column(df, ["Start Date", "Open Date", "Opened Date",
                                   "Created Date", "Create Date",
                                   "Claim Open Date", "Date Opened"])
    cd_col     = find_column(df, ["Completed Date", "Close Date", "Closed Date"])
    task_col   = find_column(df, ["Task Name", "Task"])
    labels_col = find_column(df, ["Labels", "Label"])
    due_col    = find_column(df, ["Due Date", "Due date", "Task Due", "Due"])

    if wh_col is None:
        raise KeyError("Warehouse column not found — checked: Bucket Name, Warehouse, WH, Facility")
    if sd_col is None:
        raise KeyError("Start/Open date column not found — checked: Start Date, Open Date, Created Date, etc.")

    out = df.copy()
    rename_map = {wh_col: "Warehouse", sd_col: "Start Date"}
    out.rename(columns=rename_map, inplace=True)

    if cd_col is not None:
        out.rename(columns={cd_col: "Completed Date"}, inplace=True)
        rename_map[cd_col] = "Completed Date"
    else:
        out["Completed Date"] = pd.NaT

    if task_col is not None:
        out.rename(columns={task_col: "Task Name"}, inplace=True)
        rename_map[task_col] = "Task Name"
    else:
        out["Task Name"] = ""

    if labels_col is not None:
        out.rename(columns={labels_col: "Labels"}, inplace=True)
        rename_map[labels_col] = "Labels"
    else:
        out["Labels"] = ""

    if due_col is not None:
        out.rename(columns={due_col: "Due Date"}, inplace=True)
        rename_map[due_col] = "Due Date"
    else:
        out["Due Date"] = pd.NaT

    # Type casting
    out["Start Date"]      = pd.to_datetime(out["Start Date"],      errors="coerce")
    out["Completed Date"]  = pd.to_datetime(out["Completed Date"],  errors="coerce")
    out["Due Date"]        = pd.to_datetime(out["Due Date"],         errors="coerce")
    out["Warehouse"]       = out["Warehouse"].astype(str).str.strip()
    out["Task Name"]       = out["Task Name"].astype(str).str.strip()
    out["Labels"]          = out["Labels"].astype(str)

    return out, rename_map
