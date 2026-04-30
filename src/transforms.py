# src/transforms.py — Business logic, KPI calculations, and validation.
# No I/O here. Notebooks call these functions; they never contain this logic directly.
import re
import logging
from typing import Optional

import numpy as np
import pandas as pd

logger = logging.getLogger(__name__)

BUCKET_LABELS = ["<30 days", "30-<60 days", "60-<90 days", ">= 90 days"]
BUCKET_BINS   = [-1, 29, 59, 89, float("inf")]


# ---------------------------------------------------------------------------
# Filtering
# ---------------------------------------------------------------------------

def filter_by_warehouses_and_dates(
    df: pd.DataFrame,
    warehouses: list[str],
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


# ---------------------------------------------------------------------------
# KPI calculations
# ---------------------------------------------------------------------------

def closed_claims_summary(df: pd.DataFrame, dataset_name: str) -> pd.DataFrame:
    """
    Compute closed-claims statistics (count, mean/median/min/max days-to-close)
    grouped by Warehouse.
    """
    closed = df[df["Completed Date"].notna()].copy()
    closed["DaysToClose"] = (closed["Completed Date"] - closed["Start Date"]).dt.days

    agg = (
        closed.groupby("Warehouse", dropna=False)["DaysToClose"]
        .agg(
            CountClosed="count",
            MeanDaysToClose=lambda s: float(np.nanmean(s))   if len(s) else np.nan,
            MedianDaysToClose=lambda s: float(np.nanmedian(s)) if len(s) else np.nan,
            MinDaysToClose=lambda s: float(np.nanmin(s))    if len(s) else np.nan,
            MaxDaysToClose=lambda s: float(np.nanmax(s))    if len(s) else np.nan,
        )
        .reset_index()
    )
    agg.insert(0, "Dataset", dataset_name)
    return agg


def open_claims_aging_table(
    df: pd.DataFrame,
    dataset_name: str,
    warehouses: list[str],
) -> pd.DataFrame:
    """
    Build open-claims aging buckets (<30 / 30-<60 / 60-<90 / >=90 days) per warehouse.
    Returns a full cross-product (warehouse × bucket) with Count, Percent, CumPercent.
    """
    open_df = df[df["Completed Date"].isna()].copy()
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
        .rename("Count")
        .reset_index()
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

    counts["Percent"] = counts.groupby("Warehouse")["Count"].transform(
        lambda x: (x / x.sum() * 100.0) if x.sum() else 0.0
    )
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


# ---------------------------------------------------------------------------
# Open-task extraction
# ---------------------------------------------------------------------------

def extract_wh_actions(labels_value: str) -> list[str]:
    """
    Parse a semicolon/pipe/comma-delimited Labels string and return only
    the entries that start with 'WH' (warehouse action required tags).
    """
    parts = re.split(r"[;|,]", str(labels_value))
    return [p.strip() for p in parts if re.match(r"(?i)^WH\b", p.strip())]


def build_open_tasks(
    df_std_f: pd.DataFrame,
    dataset_name: str,
) -> pd.DataFrame:
    """
    Explode WH-action labels for open claims into one row per action.
    Returns columns: Dataset, Warehouse, Task Name, Action Required, Due Date.
    """
    open_df = df_std_f[df_std_f["Completed Date"].isna()].copy()
    open_df["ActionList"] = open_df["Labels"].apply(extract_wh_actions)
    open_df = open_df[open_df["ActionList"].map(len) > 0].copy()
    open_df = open_df.explode("ActionList")
    open_df.rename(columns={"ActionList": "Action Required"}, inplace=True)
    open_df["Dataset"] = dataset_name
    return open_df[["Dataset", "Warehouse", "Task Name", "Action Required", "Due Date"]]
