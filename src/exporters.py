# src/exporters.py — All output writers: charts, Excel workbooks, CSV.
# No business logic here. Called from notebook Step 7 (Visualize) and Step 8 (Export).
from datetime import date
from pathlib import Path

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.ticker import MaxNLocator, PercentFormatter
import pandas as pd

from src.loaders import safe_filename
from src.transforms import BUCKET_LABELS


# ---------------------------------------------------------------------------
# Chart exports
# ---------------------------------------------------------------------------

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
    labels   = df_aging_wh["Aging Bucket"].astype(str).tolist()
    x        = np.arange(len(labels))
    counts   = df_aging_wh["Count"].to_numpy(dtype=float)
    cum_pct  = df_aging_wh["CumPercent"].to_numpy(dtype=float)

    fig, ax = plt.subplots(figsize=(8, 4.6))
    bars = ax.bar(x, counts)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=0)
    ax.set_xlabel("Claim Age")
    ax.set_ylabel("Open Claims (count)")
    ax.yaxis.set_major_locator(MaxNLocator(integer=True))

    ax2 = ax.twinx()
    ax2.plot(x, cum_pct, marker="o", color="tab:orange")
    ax2.set_ylabel("Cumulative %")
    ax2.set_ylim(0, 100)
    ax2.yaxis.set_major_formatter(PercentFormatter(xmax=100))

    total = int(np.nansum(counts))
    ax.set_title(f"{dataset_name} — Open Claims Aging — {wh} (n={total})")
    ax.grid(axis="y", linestyle="--", alpha=0.3)
    fig.tight_layout()

    if len(labels) <= max_labels:
        for rect, val in zip(bars, counts):
            ax.text(
                rect.get_x() + rect.get_width() / 2,
                rect.get_height(),
                f"{int(val)}",
                ha="center", va="bottom", fontsize=9,
            )

    charts_path.mkdir(parents=True, exist_ok=True)
    fname = f"{safe_filename(dataset_name)}__aging_pareto__{safe_filename(wh)}.png"
    out_path = charts_path / fname
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return out_path


# ---------------------------------------------------------------------------
# Excel exports
# ---------------------------------------------------------------------------

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
        All_Claims_Combined
        Closed_Claims_Summary
        Open_Claims_Aging
        Open_Aging_Summary
        Open_Tasks_By_Warehouse
        Open_Tasks_Aggregate
    """
    exports_path.mkdir(parents=True, exist_ok=True)
    out_path = exports_path / filename

    with pd.ExcelWriter(str(out_path), engine="openpyxl") as writer:
        if not combined_filtered.empty:
            combined_filtered.to_excel(writer, sheet_name="All_Claims_Combined", index=False)
        if not closed_all.empty:
            closed_all.to_excel(writer, sheet_name="Closed_Claims_Summary", index=False)
        if not aging_all.empty:
            aging_all.to_excel(writer, sheet_name="Open_Claims_Aging", index=False)
        if not aging_summary.empty:
            aging_summary.to_excel(writer, sheet_name="Open_Aging_Summary", index=False)
        if not open_tasks_detail.empty:
            open_tasks_detail.to_excel(writer, sheet_name="Open_Tasks_By_Warehouse", index=False)
        if not open_tasks_agg.empty:
            open_tasks_agg.to_excel(writer, sheet_name="Open_Tasks_Aggregate", index=False)

    return out_path
