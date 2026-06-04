# src/exporters.py — All output writers: charts, Excel workbooks.
# No business logic here. Called from notebook Step 7 (Visualize) and Step 8 (Export).
import logging
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.ticker import MaxNLocator, PercentFormatter

from src.transforms import BUCKET_LABELS

log = logging.getLogger(__name__)


# ── Helpers ────────────────────────────────────────────────────────────────

def safe_filename(name: str, max_len: int = 180) -> str:
    """Sanitise a string for use as a filename component."""
    return re.sub(r"[^\w\-]+", "_", str(name))[:max_len].strip("_")


# ── Excel exports ──────────────────────────────────────────────────────────

def export_to_excel(sheets: Dict[str, pd.DataFrame], output_path: Path) -> None:
    """
    Write a dict of {sheet_name: DataFrame} to a single Excel workbook.
    Skips empty DataFrames silently. Creates parent directories if needed.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                log.info("  → %-30s  %d rows", sheet_name, len(df))
    log.info("Workbook saved: %s", output_path.name)


def export_claims_workbook(
    combined_filtered: pd.DataFrame,
    closed_all: pd.DataFrame,
    aging_all: pd.DataFrame,
    aging_summary: pd.DataFrame,
    open_tasks_detail: pd.DataFrame,
    open_tasks_agg: pd.DataFrame,
    exports_path: Path,
    filename: str,
) -> Path:
    """
    Write all pipeline outputs to a single multi-sheet Excel workbook.
    Overwrites on each run to prevent stale data accumulation.

    Sheets written (skipped if DataFrame is empty):
        All_Claims_Combined, Closed_Claims_Summary, Open_Claims_Aging,
        Open_Aging_Summary, Open_Tasks_By_Warehouse, Open_Tasks_Aggregate
    """
    exports_path.mkdir(parents=True, exist_ok=True)
    out_path = exports_path / filename

    sheets = {
        "All_Claims_Combined":    combined_filtered,
        "Closed_Claims_Summary":  closed_all,
        "Open_Claims_Aging":      aging_all,
        "Open_Aging_Summary":     aging_summary,
        "Open_Tasks_By_Warehouse": open_tasks_detail,
        "Open_Tasks_Aggregate":   open_tasks_agg,
    }
    export_to_excel(sheets, out_path)
    return out_path


# ── Pareto charts ──────────────────────────────────────────────────────────

def plot_pareto_for_wh(
    df_aging_wh: pd.DataFrame,
    dataset_name: str,
    wh: str,
    max_labels: int,
    charts_path: Path,
) -> Path:
    """
    Render and save a Pareto chart (bar + cumulative % line) for one
    warehouse's open-claims aging distribution.

    Returns the saved file path.
    """
    total   = int(df_aging_wh["Count"].sum())
    labels  = df_aging_wh["Aging Bucket"].astype(str).tolist()
    counts  = df_aging_wh["Count"].tolist()
    cum_pct = df_aging_wh["CumPercent"].tolist()

    fig, ax1 = plt.subplots(figsize=(8, 4.6))
    ax2 = ax1.twinx()

    bars = ax1.bar(labels, counts)
    ax2.plot(labels, cum_pct, marker="o", color="tab:orange", zorder=5)

    ax1.set_xlabel("Claim Age")
    ax1.set_ylabel("Open Claims (count)")
    ax1.yaxis.set_major_locator(MaxNLocator(integer=True))
    ax1.yaxis.grid(True, linestyle="--", alpha=0.3)
    ax1.set_axisbelow(True)

    ax2.set_ylabel("Cumulative %")
    ax2.set_ylim(0, 100)
    ax2.yaxis.set_major_formatter(PercentFormatter(xmax=100))

    if len(labels) <= max_labels:
        for bar, val in zip(bars, counts):
            if val > 0:
                ax1.text(
                    bar.get_x() + bar.get_width() / 2,
                    bar.get_height() + 0.1,
                    str(val),
                    ha="center", va="bottom", fontsize=9,
                )

    ax1.set_title(f"{dataset_name} — Open Claims Aging — {wh} (n={total})")
    plt.xticks(rotation=0)
    plt.tight_layout()

    charts_path.mkdir(parents=True, exist_ok=True)
    fname = f"{safe_filename(dataset_name)}__aging_pareto__{safe_filename(wh)}.png"
    out_path = charts_path / fname
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    log.info("  Chart saved: %s", out_path.name)
    return out_path


def generate_all_pareto_charts(
    aging_all: pd.DataFrame,
    fig_dir: Path,
    cfg: dict,
) -> None:
    """
    Iterate all dataset × warehouse combinations and call plot_pareto_for_wh.
    """
    desired_order = BUCKET_LABELS
    fig_dir.mkdir(parents=True, exist_ok=True)

    for dataset_name in aging_all["Dataset"].unique():
        df_ds = aging_all.loc[aging_all["Dataset"] == dataset_name]
        for wh in df_ds["Warehouse"].unique():
            df_wh = (
                df_ds.loc[df_ds["Warehouse"] == wh]
                .set_index("Aging Bucket")
                .reindex(desired_order)
                .fillna(0)
                .reset_index()
            )
            plot_pareto_for_wh(
                df_wh,
                dataset_name,
                wh,
                max_labels=cfg.get("max_labels", 25),
                charts_path=fig_dir,
            )

    log.info("All Pareto charts saved → %s/", fig_dir.name)
