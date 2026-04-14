"""
src/transforms.py
-----------------
All business logic: filtering, KPI aggregations, aging bucketing, and
open-task/tag extraction. No I/O here — pure DataFrame → DataFrame.
"""
import logging
import re
from typing import List, Optional, Tuple

import numpy as np
import pandas as pd

log = logging.getLogger(__name__)

_WH_ACTION_PATTERN = re.compile(r"\bwh\s*please\s*respo?nd\b", re.IGNORECASE)


# ── Filtering ──────────────────────────────────────────────────────────────

def filter_by_warehouses_and_dates(
    df: pd.DataFrame,
    warehouses: List[str],
    start_date: str,
    end_date: str,
) -> pd.DataFrame:
    """
    Filter DataFrame to the specified warehouses and Start Date range.

    Warehouse matching is case-insensitive; output values are normalised
    to the case provided in the warehouses list.
    """
    wh_lower = {w.lower(): w for w in warehouses}
    mask_wh  = df["Warehouse"].str.lower().isin(wh_lower)
    mask_dt  = df["Start Date"].between(
        pd.Timestamp(start_date), pd.Timestamp(end_date)
    )
    df_f = df.loc[mask_wh & mask_dt].copy()
    df_f["Warehouse"] = df_f["Warehouse"].str.lower().map(wh_lower)
    log.info("  filter → %d rows", len(df_f))
    return df_f


# ── Closed claims KPIs ─────────────────────────────────────────────────────

def closed_claims_summary(df: pd.DataFrame, dataset_name: str) -> pd.DataFrame:
    """
    Aggregate closed-claim cycle-time KPIs by warehouse.

    Returns columns: Dataset, Warehouse, CountClosed,
                     MeanDaysToClose, MedianDaysToClose,
                     MinDaysToClose, MaxDaysToClose
    """
    closed = df.loc[df["Completed Date"].notna()].copy()
    closed["DaysToClose"] = (closed["Completed Date"] - closed["Start Date"]).dt.days

    summary = (
        closed.groupby("Warehouse")["DaysToClose"]
        .agg(
            CountClosed="count",
            MeanDaysToClose="mean",
            MedianDaysToClose="median",
            MinDaysToClose="min",
            MaxDaysToClose="max",
        )
        .reset_index()
    )
    summary.insert(0, "Dataset", dataset_name)
    return summary


# ── Open claims aging ──────────────────────────────────────────────────────

def open_claims_aging_table(
    df: pd.DataFrame,
    dataset_name: str,
    warehouses: List[str],
    bins: List,
    labels: List[str],
) -> pd.DataFrame:
    """
    Build open-claims aging distribution table.

    Returns columns: Dataset, Warehouse, Aging Bucket, Count, Percent, CumPercent
    """
    open_df = df.loc[df["Completed Date"].isna()].copy()
    today   = pd.Timestamp("today").normalize()
    open_df["AgeDays"] = (today - open_df["Start Date"]).dt.days

    open_df["Aging Bucket"] = pd.cut(
        open_df["AgeDays"], bins=bins, labels=labels, right=True
    )

    counts = (
        open_df.groupby(["Warehouse", "Aging Bucket"], observed=False)
        .size()
        .reset_index(name="Count")
    )

    # Ensure every warehouse × bucket combination is represented
    full_idx = pd.MultiIndex.from_product(
        [warehouses, labels], names=["Warehouse", "Aging Bucket"]
    )
    counts = (
        counts.set_index(["Warehouse", "Aging Bucket"])
        .reindex(full_idx, fill_value=0)
        .reset_index()
    )

    totals = counts.groupby("Warehouse")["Count"].transform("sum")
    counts["Percent"]    = (counts["Count"] / totals.replace(0, np.nan) * 100).fillna(0)
    counts["CumPercent"] = counts.groupby("Warehouse")["Percent"].cumsum()
    counts.insert(0, "Dataset", dataset_name)
    return counts


def build_aging_summary(aging_all: pd.DataFrame, aging_labels: List[str]) -> pd.DataFrame:
    """
    Pivot aging table into a leadership-ready summary.

    Returns: Dataset, Warehouse, <bucket cols>, Total Open Claims, % >= 90 days
    """
    summary = (
        aging_all
        .pivot_table(
            index=["Dataset", "Warehouse"],
            columns="Aging Bucket",
            values="Count",
            aggfunc="sum",
        )
        .reindex(columns=aging_labels, fill_value=0)
        .reset_index()
    )
    summary["Total Open Claims"] = summary[aging_labels].sum(axis=1)
    summary["% >= 90 days"] = (
        summary[">= 90 days"] / summary["Total Open Claims"].replace(0, np.nan)
    ).fillna(0)
    return summary


# ── Open tasks ─────────────────────────────────────────────────────────────

def _extract_wh_actions(labels_value: Optional[str]) -> List[str]:
    """Parse a Labels string and return items that start with 'WH'."""
    if not isinstance(labels_value, str) or not labels_value.strip():
        return []
    parts = re.split(r"[;|,]", labels_value)
    return [p.strip() for p in parts if p.strip().upper().startswith("WH")]


def build_open_tasks(df: pd.DataFrame, dataset_name: str) -> Optional[pd.DataFrame]:
    """
    Extract open claims with warehouse action labels (exploded to one row per action).

    Returns columns: Dataset, Warehouse, Task Name, Action Required, Due Date
    Returns None if no matching rows found.
    """
    if "Labels" not in df.columns:
        return None

    open_df = df.loc[df["Completed Date"].isna()].copy()
    open_df["ActionList"] = open_df["Labels"].apply(_extract_wh_actions)
    tasks = open_df.loc[open_df["ActionList"].map(len) > 0].copy()

    if tasks.empty:
        return None

    tasks = tasks.explode("ActionList").rename(columns={"ActionList": "Action Required"})
    keep  = [c for c in ["Dataset", "Warehouse", "Task Name", "Action Required", "Due Date"] if c in tasks.columns]
    return tasks[keep].reset_index(drop=True)


# ── WH tag scan ────────────────────────────────────────────────────────────

def scan_wh_tags(df: pd.DataFrame, dataset_name: str) -> Optional[pd.DataFrame]:
    """
    Scan Labels column for 'WH Please Respond' pattern (regex, full dataset).

    Returns: Dataset, Warehouse, Task Name — deduplicated by Task Name.
    Returns None if Labels column absent or no matches.
    """
    if "Labels" not in df.columns:
        return None

    df_norm = df.copy()
    mask    = df_norm["Labels"].str.contains(_WH_ACTION_PATTERN, na=False)
    hits    = df_norm.loc[mask, [c for c in ["Warehouse", "Task Name", "Labels"] if c in df_norm.columns]].copy()

    if hits.empty:
        return None

    hits.insert(0, "Dataset", dataset_name)
    hits = hits.drop_duplicates(subset=["Task Name"]).sort_values(["Warehouse", "Task Name"])
    log.info("  WH tags found: %d (dataset=%s)", len(hits), dataset_name)
    return hits.reset_index(drop=True)
