# src/transforms.py — Business logic, KPI calculations, and validation.
# No I/O here. Notebooks call these functions; they never contain this logic directly.
import logging
import re
from typing import List, Optional

import numpy as np
import pandas as pd

log = logging.getLogger(__name__)

BUCKET_LABELS = ["<30 days", "30-<60 days", "60-<90 days", ">= 90 days"]
BUCKET_BINS   = [-1, 29, 59, 89, float("inf")]

_WH_ACTION_PATTERN = re.compile(r"\bwh\s*please\s*respo?nd\b", re.IGNORECASE)


# ── Filtering ──────────────────────────────────────────────────────────────

def filter_by_warehouses_and_dates(
    df: pd.DataFrame,
    warehouses: List[str],
    start_date: Optional[str],
    end_date: Optional[str],
) -> pd.DataFrame:
    """
    Filter to the specified warehouses and date range (inclusive on both ends).
    Warehouse matching is case-insensitive; original casing is restored from the
    warehouses list.
    """
    wanted = {w.strip().lower(): w for w in warehouses}
    df2 = df[df["Warehouse"].str.lower().isin(wanted.keys())].copy()
    df2["Warehouse"] = df2["Warehouse"].str.lower().map(wanted)

    if start_date:
        df2 = df2[df2["Start Date"] >= pd.to_datetime(start_date)]
    if end_date:
        df2 = df2[df2["Start Date"] <= pd.to_datetime(end_date)]
    return df2


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
) -> pd.DataFrame:
    """
    Build open-claims aging buckets (<30 / 30-<60 / 60-<90 / >=90 days) per warehouse.
    Returns a full cross-product (warehouse × bucket) with Count, Percent, CumPercent.
    """
    open_df = df.loc[df["Completed Date"].isna()].copy()
    today = pd.Timestamp.today().normalize()
    open_df["AgeDays"] = (today - open_df["Start Date"].dt.normalize()).dt.days

    open_df["Aging Bucket"] = pd.cut(
        open_df["AgeDays"],
        bins=BUCKET_BINS,
        labels=BUCKET_LABELS,
        include_lowest=True,
        right=True,
    )

    counts = (
        open_df.groupby(["Warehouse", "Aging Bucket"], observed=False)
        .size()
        .reset_index(name="Count")
    )

    # Ensure every warehouse × bucket combination is present
    active_wh = [w for w in warehouses if w in counts["Warehouse"].unique().tolist()]
    full_index = pd.MultiIndex.from_product(
        [active_wh, BUCKET_LABELS], names=["Warehouse", "Aging Bucket"]
    )
    counts = (
        counts.set_index(["Warehouse", "Aging Bucket"])
        .reindex(full_index, fill_value=0)
        .reset_index()
    )

    totals = counts.groupby("Warehouse")["Count"].transform("sum")
    counts["Percent"]    = (counts["Count"] / totals.replace(0, np.nan) * 100).fillna(0)
    counts["CumPercent"] = counts.groupby("Warehouse")["Percent"].cumsum()
    counts.insert(0, "Dataset", dataset_name)
    return counts


def build_aging_summary(aging_all: pd.DataFrame) -> pd.DataFrame:
    """
    Pivot open-claims aging into a leadership-friendly summary table
    (Dataset × Warehouse rows, bucket columns, Total Open Claims, % >= 90 days).
    """
    aging_summary = (
        aging_all
        .pivot_table(
            index=["Dataset", "Warehouse"],
            columns="Aging Bucket",
            values="Count",
            aggfunc="sum",
            fill_value=0,
        )
        .reindex(columns=BUCKET_LABELS, fill_value=0)
        .reset_index()
    )
    aging_summary.columns.name = None
    aging_summary["Total Open Claims"] = aging_summary[BUCKET_LABELS].sum(axis=1)
    aging_summary["% >= 90 days"] = np.where(
        aging_summary["Total Open Claims"] > 0,
        aging_summary[">= 90 days"] / aging_summary["Total Open Claims"],
        np.nan,
    )
    return aging_summary


# ── Open tasks ─────────────────────────────────────────────────────────────

def _extract_wh_actions(labels_value: Optional[str]) -> List[str]:
    """Parse a Labels string and return items that start with 'WH'."""
    if not isinstance(labels_value, str) or not labels_value.strip():
        return []
    parts = re.split(r"[;|,]", labels_value)
    return [p.strip() for p in parts if p.strip().upper().startswith("WH")]


def build_open_tasks(
    df_std_f: pd.DataFrame,
    dataset_name: str,
) -> pd.DataFrame:
    """
    Explode WH-action labels for open claims into one row per action.
    Returns columns: Dataset, Warehouse, Task Name, Action Required, Due Date.
    """
    _empty = pd.DataFrame(
        columns=["Dataset", "Warehouse", "Task Name", "Action Required", "Due Date"]
    )
    if "Labels" not in df_std_f.columns:
        return _empty

    open_df = df_std_f.loc[df_std_f["Completed Date"].isna()].copy()
    open_df["ActionList"] = open_df["Labels"].apply(_extract_wh_actions)
    open_df = open_df[open_df["ActionList"].map(len) > 0].copy()
    if open_df.empty:
        return _empty

    open_df = open_df.explode("ActionList")
    open_df.rename(columns={"ActionList": "Action Required"}, inplace=True)
    open_df["Dataset"] = dataset_name
    return open_df[["Dataset", "Warehouse", "Task Name", "Action Required", "Due Date"]]


# ── WH tag scan ────────────────────────────────────────────────────────────

def scan_wh_tags(df: pd.DataFrame, dataset_name: str) -> Optional[pd.DataFrame]:
    """
    Scan Labels column for 'WH Please Respond' pattern (full dataset, unfiltered).
    Returns: Dataset, Warehouse, Task Name — deduplicated by Task Name.
    Returns None if Labels column absent or no matches.
    """
    if "Labels" not in df.columns:
        return None

    mask = df["Labels"].str.contains(_WH_ACTION_PATTERN, na=False)
    hits = df.loc[mask, [c for c in ["Warehouse", "Task Name", "Labels"] if c in df.columns]].copy()

    if hits.empty:
        return None

    hits.insert(0, "Dataset", dataset_name)
    hits = hits.drop_duplicates(subset=["Task Name"]).sort_values(["Warehouse", "Task Name"])
    log.info("  WH tags found: %d (dataset=%s)", len(hits), dataset_name)
    return hits.reset_index(drop=True)
