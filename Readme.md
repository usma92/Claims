# Claims Aging Analysis

Pipeline for tracking open Ford and Chrysler dealer claims by warehouse and age bucket.
Pulls data monthly from Microsoft Planner via Graph API and produces an HTML aging dashboard.

---

## What it produces

| Output | Location | Refreshed |
|---|---|---|
| HTML dashboard (latest) | `dashboard.html` | Every run |
| Dated HTML dashboard | `dashboards/YYYYMMDD_claims_aging_dashboard.html` | Every run |
| Excel workbook | `claims_analysis_output.xlsx` | Every run |
| Charts | `charts/` | Every run |

The dashboard shows open claim counts by age bucket (<30 / 30-60 / 60-90 / ≥90 days)
per warehouse, for Ford and Chrysler separately.

---

## Project Structure

```
2. Claims/
├── notebooks/
│   └── 00_claims_aging_pipeline.ipynb   # Main pipeline — run this
├── src/
│   ├── paths.py        # Dynamic path resolution (P = get_paths())
│   ├── loaders.py      # Excel ingest + column normalization
│   ├── transforms.py   # Aging calculations, open-claim filtering
│   └── exporters.py    # Charts, Excel, HTML dashboard
├── scripts/
│   ├── download_planner.py         # Graph API download (MSAL device flow)
│   ├── download_planner_browser.py # Playwright fallback downloader
│   └── run_pipeline.bat            # Task Scheduler entry point
├── config/
│   └── settings.yaml   # Input files, warehouses, aging bins, date range
├── data/
│   ├── raw/            # Ford_Claims.xlsx, Chrysler_Claims.xlsx (git-ignored)
│   └── exports/        # Dated Excel output (git-ignored)
├── dashboards/         # Dated HTML dashboards (git-ignored)
├── charts/             # PNG charts (git-ignored)
├── SETUP_AUTH.md       # One-time Azure AD app registration guide
└── requirements.txt
```

---

## Data Source

Claims are stored in Microsoft Planner (two plans: Ford Claims, Chrysler Claims).
`scripts/download_planner.py` calls the Graph API using MSAL device-flow auth and
writes `data/raw/Ford_Claims.xlsx` and `data/raw/Chrysler_Claims.xlsx`.

**First-time setup:** See `SETUP_AUTH.md` — requires an Azure AD app registration
(CLIENT_ID) and one interactive browser login to cache the refresh token.

**Subsequent runs:** Fully silent — cached token handles auth with no browser prompt.

---

## Running the Pipeline

### Step 1 — Download fresh data (monthly)

```powershell
cd "C:\Users\dbalan\Documents\Claude\Projects\Team Dashboards\2. Claims"
.venv\Scripts\activate.bat
python scripts\download_planner.py
```

### Step 2 — Run the notebook

Open `notebooks/00_claims_aging_pipeline.ipynb` in Jupyter, then:

```
Kernel → Restart Kernel and Run All Cells
```

Or headless via the publisher (`run_all.py` handles this automatically on its monthly schedule).

---

## Configuration (`config/settings.yaml`)

| Key | Description |
|---|---|
| `input_files` | List of `{path, name}` — paths relative to `data/raw/` |
| `warehouses` | Warehouse names to include in output |
| `start_date` | Earliest claim start date to include (`null` = no filter) |
| `end_date` | Latest date (`null` = today) |
| `aging_bins` | Day breakpoints for age buckets |
| `aging_labels` | Labels for each bucket |

---

## Column Schema

After normalization (`src/loaders.py`), all input files share this schema:

| Column | Source | Notes |
|---|---|---|
| `Warehouse` | `Bucket Name` or `bucket` | Location identifier |
| `Start Date` | `Created Date` / `open date` / etc. | Claim open date |
| `Completed Date` | `close date` / `resolved date` / etc. | Blank = still open |
| `Task Name` | `task` / `claim` / `description` | Claim description |
| `Labels` | `labels` / `tags` | Semicolon-delimited issue tags |

An open claim is any row where `Completed Date` is NaT.

---

## Publisher Integration

Registered in `_publisher/config.yaml` with `runner: "nbconvert:notebooks/00_claims_aging_pipeline.ipynb"`,
`run: true`, and `freshness_hours: 840` (monthly cadence — ~35 days before yellow badge).
The publisher picks up `dashboard.html` from the project root.
