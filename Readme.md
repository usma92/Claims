# Claims Performance Analysis (Multi‑File) — README

This project analyzes **dealer claims** from one or more Excel workbooks and produces:
- **Cycle‑time performance by Bucket Name** (count, mean, median, min, max days)
- **Pareto of Labels** (overall and **by warehouse**)
- **Power BI–ready exports** for dashboards
- A **date‑range filter** that can be applied to *Created Date*, *Completed Date*, or *Either*

Primary notebook: **`Claims_Analysis_Pareto_Multi.ipynb`**

---

## Inputs & Assumptions

### Required files
Place your Excel workbooks in the same folder as the notebook (or update the paths):
- `Claims.xlsx`
- `Chrysler Claims.xlsx`  
*(You can add more; see **Parameters** below.)*

### Required sheet & columns
- Sheet: **`Tasks`**
- Columns: **`Bucket Name`**, **`Labels`**, **`Created Date`**, **`Completed Date`**
  - `Labels` may contain multiple labels separated by `;` (e.g., `DD/DB Shortage; Outbound freight damage`).

### Date filter (inclusive)
You can limit analysis to a date window by **Created Date**, **Completed Date**, or **Either** (rows where *either* date falls in range).

---

## Outputs

A single Excel workbook is written to disk (default: **`Claims_Performance_Summary_ALL.xlsx`**), containing Power BI–ready tables:

- `Bucket_Performance` — **Dataset, Bucket Name, Count, Mean, Median, Min, Max**
- `Bucket_Performance_ALL` — Combined view across *all* datasets
- `Pareto_Labels_Overall` — (ALL + per‑dataset) **Label, Count, Cumulative %**
- `Pareto_Labels_By_Warehouse` — (ALL + per‑dataset) **Warehouse, Label, Count, Percent, Cumulative %**
- `Detailed_Data` — Underlying filtered rows with **Cycle Time (Days)** and **Dataset**

> **Cycle Time (Days) = Completed Date − Created Date.**

---

## Quick Start

1. Open **`Claims_Analysis_Pareto_Multi.ipynb`** in Jupyter (VS Code, JupyterLab, etc.).  
2. In the **Parameters** cell, confirm file names, warehouse list, and date filter settings.  
3. Run all cells.  
4. Find the export at **`Claims_Performance_Summary_ALL.xlsx`** and connect it to Power BI (Get Data → Excel).

### Dependencies
Python 3.9+ recommended.
```bash
pip install pandas matplotlib openpyxl
```

---

## Parameters (edit in the notebook)

```python
INPUT_FILES = [
    {"name": "Claims", "filename": "Claims.xlsx"},
    {"name": "Chrysler Claims", "filename": "Chrysler Claims.xlsx"},
]

SHEET_NAME  = "Tasks"
OUTPUT_XLSX = Path("Claims_Performance_Summary_ALL.xlsx")

# Warehouses used for the by‑warehouse Pareto (filter on Bucket Name)
WAREHOUSES  = ["OKC", "Atlanta", "Orlando", "Ontario", "El Paso", "Flowood", "Phoenix", "Charlotte"]

# Max labels to show in each Pareto chart (set None to show all)
TOP_N_LABELS = 15

# Date Range Filter (inclusive)
DATE_FILTER_MODE = "Created"  # "Created", "Completed", or "Either"
START_DATE = None             # e.g., "2025-01-01"
END_DATE   = None             # e.g., "2025-06-30"
```

---

## Cell‑by‑Cell Explanation

**1) Markdown — Title & Scope**  
States the notebook’s purpose: multi‑file analysis, cycle times, Pareto (overall & by warehouse), Power BI exports, and date filtering.

**2) Markdown — Inputs**  
Describes expected files/sheet and required columns.

**3) Code — Imports & Display Settings**  
Imports `Path`, `numpy`, `pandas`, `matplotlib.pyplot`; widens Pandas display for readability in previews.

**4) Code — Parameters**  
Central place to configure:
- Which files to load (`INPUT_FILES`)
- Sheet name (`SHEET_NAME`)
- Output Excel file path (`OUTPUT_XLSX`)
- Warehouse list for by‑warehouse Pareto (`WAREHOUSES`)
- Max labels per Pareto chart (`TOP_N_LABELS`)
- **Date filter** mode and window (`DATE_FILTER_MODE`, `START_DATE`, `END_DATE`)

**5) Code — Helper Functions**
- `resolve_path(fname)`: Finds each workbook (working dir → `/mnt/data` fallback).
- `load_and_prepare(path, sheet, dataset_name)`: Loads `Tasks`, validates columns, strips text, parses dates, computes **Cycle Time (Days)**, tags the source `Dataset` name.
- `apply_date_filter(df, mode, start, end)`: Returns only rows falling **within** `[start, end]` by the chosen mode (**Created**, **Completed**, or **Either**).
- `explode_labels(df)`: Splits semicolon‑separated labels into long format (one row per label); trims and drops empties.
- `compute_bucket_perf(df)`: By **Dataset + Bucket Name**, returns **Count, Mean, Median, Min, Max** of cycle time (rounded to 2 decimals).
- `compute_overall_pareto(labels_exploded)`: Overall label frequency across **ALL** datasets with **Cumulative %**, tagged `Dataset="ALL"`.
- `compute_by_wh_pareto(labels_exploded, warehouses)`: By **Dataset + Warehouse**, returns **Count, Percent, Cumulative %**, sorted by count.
- `plot_pareto(sub, title, top_n)`: Makes a single Pareto chart: **Count** bars + **Cumulative %** line (0–100%), one figure per warehouse/dataset.

**6) Code — Load & Combine Datasets**  
Loops over `INPUT_FILES`, loads each with `load_and_prepare`, prints status, concatenates into `df_all`. Errors are reported but don’t stop the whole run unless nothing loads.

**7) Code — Apply Date Filter**  
Creates `df_filtered = apply_date_filter(df_all, DATE_FILTER_MODE, START_DATE, END_DATE)`.  
Prints before/after row counts and shows a small preview.

**8) Code — Bucket Performance (Filtered)**  
- `bucket_perf_by_dataset`: Per **Dataset + Bucket Name** stats.  
- `bucket_perf_all`: Same stats aggregated for the combined **ALL** dataset.  
Shows the head of the dataset‑level table.

**9) Code — Bucket Performance (ALL) Preview**  
Displays the head of `bucket_perf_all` (combined table).

**10) Code — Pareto Calculations (Filtered)**  
- `labels_exploded`: Long format labels from filtered data.  
- `overall_all`: Overall counts across ALL datasets with **Cumulative %**.  
- `overall_by_dataset`: Per‑dataset counts with per‑dataset **Cumulative %**.  
- `by_wh`: Per‑dataset **Warehouse/Label** counts with **Percent** and **Cumulative %**.  
- `by_wh_all`: Combined **ALL** per‑warehouse table.  
Shows the first 15 rows of `overall_all`.

**11) Code — Pareto Charts**  
Generates Pareto charts for each **Dataset + Warehouse** (limited to `TOP_N_LABELS`), then repeats for the combined **ALL** dataset.

**12) Code — Export for Power BI (Filtered)**  
Writes **five** sheets to `OUTPUT_XLSX`:  
- `Bucket_Performance` (per dataset)  
- `Bucket_Performance_ALL` (combined)  
- `Pareto_Labels_Overall` (ALL + per‑dataset)  
- `Pareto_Labels_By_Warehouse` (ALL + per‑dataset)  
- `Detailed_Data` (filtered rows with `Cycle Time (Days)` and `Dataset`)  
Prints the resolved path after export.

---

## Power BI Tips
- Use *Get Data → Excel* and connect to `Claims_Performance_Summary_ALL.xlsx`.
- Mark your date columns as **date** types.
- Popular visuals:
  - **Pareto**: Clustered column (Count) + line (Cumulative %) with shared X axis.
  - **Bucket performance**: Table or bar chart of **Mean/Median** cycle time by **Bucket Name**.
  - Add **Slicers** for `Dataset`, `Warehouse`, and `Labels`.  
- Keep your data files in a consistent path and set up a refresh schedule if desired.

---

## Troubleshooting
- **Missing columns**: Ensure `Bucket Name`, `Labels`, `Created Date`, `Completed Date` exist on the `Tasks` sheet.  
- **Empty results after filtering**: Check `DATE_FILTER_MODE`, `START_DATE`, `END_DATE`. Try `DATE_FILTER_MODE="Either"`.  
- **Too many labels on charts**: Reduce `TOP_N_LABELS`.  
- **Wrong warehouses in by‑warehouse Pareto**: Update the `WAREHOUSES` list.

---

*Maintained for operational claims analysis and reporting. For enhancements (e.g., SLA bands, open‑item views, or trend lines), extend the helper functions and exports as needed.*
