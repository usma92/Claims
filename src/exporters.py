"""
src/exporters.py
----------------
All output writers: Excel workbook export and Pareto chart generation.
No business logic here — pure rendering and I/O.
"""
import logging
import re
from pathlib import Path
from typing import Dict, List, Tuple

import matplotlib.pyplot as plt
import pandas as pd
from matplotlib.ticker import MaxNLocator, PercentFormatter

log = logging.getLogger(__name__)


# ── Helpers ────────────────────────────────────────────────────────────────

def safe_filename(name: str, max_len: int = 180) -> str:
    """Sanitise a string for use as a filename component."""
    return re.sub(r"[^\w\-]+", "_", str(name))[:max_len].strip("_")


# ── Excel export ───────────────────────────────────────────────────────────

def export_to_excel(sheets: Dict[str, pd.DataFrame], output_path: Path) -> None:
    """
    Write a dict of {sheet_name: DataFrame} to a single Excel workbook.

    Skips empty DataFrames silently. Creates parent directories if needed.

    Args:
        sheets:      Ordered dict of sheet name → DataFrame.
        output_path: Absolute path for the output .xlsx file.
    """
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                log.info("  → %-30s  %d rows", sheet_name, len(df))

    log.info("Workbook saved: %s", output_path.name)


# ── Pareto chart ───────────────────────────────────────────────────────────

def plot_pareto_for_wh(
    df_aging_wh: pd.DataFrame,
    dataset_name: str,
    wh: str,
    fig_dir: Path,
    max_labels: int = 25,
    fig_size: Tuple = (8, 4.6),
    dpi: int = 150,
) -> None:
    """
    Render and save a dual-axis Pareto chart for one dataset/warehouse pair.

    Left axis  → bar chart of open-claim counts per aging bucket.
    Right axis → cumulative % line.
    Output saved to: fig_dir/<dataset>__aging_pareto__<warehouse>.png
    """
    total   = int(df_aging_wh["Count"].sum())
    labels  = df_aging_wh["Aging Bucket"].astype(str).tolist()
    counts  = df_aging_wh["Count"].tolist()
    cum_pct = df_aging_wh["CumPercent"].tolist()

    fig, ax1 = plt.subplots(figsize=fig_size)
    ax2 = ax1.twinx()

    bars = ax1.bar(labels, counts)
    ax2.plot(labels, cum_pct, marker="o", color="tab:orange", zorder=5)

    ax1.set_xlabel("Aging Bucket")
    ax1.set_ylabel("Open Claims (count)")
    ax1.yaxis.set_major_locator(MaxNLocator(integer=True))

    ax2.set_ylabel("Cumulative %")
    ax2.set_ylim(0, 100)
    ax2.yaxis.set_major_formatter(PercentFormatter(xmax=100))

    ax1.yaxis.grid(True, linestyle="--", alpha=0.3)
    ax1.set_axisbelow(True)

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

    fname = fig_dir / f"{safe_filename(dataset_name)}__aging_pareto__{safe_filename(wh)}.png"
    fig.savefig(fname, dpi=dpi, bbox_inches="tight")
    plt.close(fig)
    log.info("  Chart saved: %s", fname.name)


def generate_all_pareto_charts(
    aging_all: pd.DataFrame,
    aging_labels: List[str],
    fig_dir: Path,
    cfg: dict,
) -> None:
    """
    Iterate all dataset × warehouse combinations and call plot_pareto_for_wh.

    Args:
        aging_all:    Concatenated open-claims aging table (all datasets).
        aging_labels: Ordered list of aging bucket labels.
        fig_dir:      Directory where PNGs are saved.
        cfg:          Settings dict (used for max_labels, fig_size, fig_dpi).
    """
    fig_dir.mkdir(parents=True, exist_ok=True)

    for dataset_name in aging_all["Dataset"].unique():
        df_ds = aging_all.loc[aging_all["Dataset"] == dataset_name]
        for wh in df_ds["Warehouse"].unique():
            df_wh = (
                df_ds.loc[df_ds["Warehouse"] == wh]
                .set_index("Aging Bucket")
                .reindex(aging_labels)
                .fillna(0)
                .reset_index()
            )
            plot_pareto_for_wh(
                df_wh,
                dataset_name,
                wh,
                fig_dir,
                max_labels=cfg.get("max_labels", 25),
                fig_size=tuple(cfg.get("fig_size", [8, 4.6])),
                dpi=cfg.get("fig_dpi", 150),
            )

    log.info("All Pareto charts saved → %s/", fig_dir.name)
