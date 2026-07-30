"""
download_planner.py — Claims data ingestion via Microsoft Graph API
====================================================================
*** MANUAL FALLBACK ONLY ***
Primary automation is the Power Automate flow "FEED — Claims (Ford + Chrysler)"
which runs on the 1st of each month and writes CSVs to data/raw/ via OneDrive.
See _publisher/POWER_AUTOMATE_RECIPES.md for the full PA flow recipe.

Use this script only if the PA flow is unavailable or you need to trigger
an out-of-cycle refresh manually. Requires an Azure AD app registration
(CLIENT_ID below) — see SETUP_AUTH.md.

Output: data/raw/Ford_Claims.xlsx and data/raw/Chrysler_Claims.xlsx
Archive: data/archive/Ford_Claims_<timestamp>.xlsx

Prerequisites: pip install msal pandas openpyxl requests
"""

from __future__ import annotations

import shutil
import sys
from datetime import datetime
from pathlib import Path

import msal
import pandas as pd
import requests

# ─── Configuration ────────────────────────────────────────────────────────────
TENANT_ID = "4455dc99-bca8-4daf-bd45-2a50a7ceb65f"

# Fill in after IT registers the Azure AD app (see SETUP_AUTH.md).
# Delegated permissions required: Tasks.Read, User.Read
CLIENT_ID = "YOUR_APP_CLIENT_ID"

SCOPES = ["Tasks.Read", "User.Read"]

PLANS = {
    "Ford":     "jRdOeWc20kuFagvcx_OkrmUABPJd",
    "Chrysler": "hYGA2uMjlE-9NFjXfoUFeWUAGOP3",
}

GRAPH = "https://graph.microsoft.com/v1.0"

HERE      = Path(__file__).resolve().parent.parent   # project root (2. Claims/)
RAW_DIR   = HERE / "data" / "raw"
ARCH_DIR  = HERE / "data" / "archive"
TOKEN_FILE = HERE / "data" / ".token_cache.json"


# ─── Auth ─────────────────────────────────────────────────────────────────────
def get_token() -> str:
    """Acquire access token via device flow; cache refresh token for silent reuse."""
    cache = msal.SerializableTokenCache()
    if TOKEN_FILE.exists():
        cache.deserialize(TOKEN_FILE.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache,
    )

    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

    if not result:
        # First run (or token expired beyond refresh window): device flow
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(f"Device flow failed: {flow}")
        print("\n" + flow["message"] + "\n")   # e.g. "Go to https://microsoft.com/devicelogin and enter XXXX-XXXX"
        result = app.acquire_token_by_device_flow(flow)

    # Persist updated cache (new access token or refreshed refresh token)
    TOKEN_FILE.write_text(cache.serialize(), encoding="utf-8")

    if "access_token" not in result:
        raise RuntimeError(f"Auth failed: {result.get('error_description', result)}")

    return result["access_token"]


# ─── Graph API helpers ────────────────────────────────────────────────────────
def graph_get_all(token: str, url: str) -> list[dict]:
    """GET a Graph endpoint, following @odata.nextLink pagination."""
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    items: list[dict] = []
    while url:
        r = requests.get(url, headers=headers, timeout=30)
        r.raise_for_status()
        data = r.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def resolve_user_names(token: str, user_ids: list[str]) -> dict[str, str]:
    """Batch-resolve AAD object IDs → display names."""
    names: dict[str, str] = {}
    for uid in set(user_ids):
        if not uid:
            continue
        try:
            r = requests.get(
                f"{GRAPH}/users/{uid}",
                headers={"Authorization": f"Bearer {token}"},
                params={"$select": "displayName"},
                timeout=10,
            )
            if r.ok:
                names[uid] = r.json().get("displayName", uid)
            else:
                names[uid] = uid  # fallback to raw ID
        except Exception:
            names[uid] = uid
    return names


# ─── Planner extraction ───────────────────────────────────────────────────────
def extract_plan(token: str, plan_id: str) -> pd.DataFrame:
    """Return a DataFrame of all tasks in a Planner plan."""
    buckets  = graph_get_all(token, f"{GRAPH}/planner/plans/{plan_id}/buckets")
    bucket_map = {b["id"]: b["name"] for b in buckets}

    tasks = graph_get_all(token, f"{GRAPH}/planner/plans/{plan_id}/tasks")

    # Collect all assigned user IDs so we can resolve names in one pass
    all_user_ids = [uid for t in tasks for uid in t.get("assignments", {}).keys()]
    user_names = resolve_user_names(token, all_user_ids)

    rows = []
    for t in tasks:
        assigned = [user_names.get(uid, uid) for uid in t.get("assignments", {}).keys()]
        rows.append({
            "Task ID":        t["id"],
            "Title":          t.get("title", ""),
            "Bucket":         bucket_map.get(t.get("bucketId", ""), ""),
            "Assigned To":    ", ".join(assigned),
            "Due Date":       (t.get("dueDateTime") or "")[:10],
            "Start Date":     (t.get("startDateTime") or "")[:10],
            "Percent Done":   t.get("percentComplete", 0),
            "Priority":       _priority_label(t.get("priority", 5)),
            "Created":        (t.get("createdDateTime") or "")[:10],
            "Has Checklist":  bool(t.get("checklistItemCount", 0)),
            "Checklist Done": t.get("activeChecklistItemCount", 0) == 0 and t.get("checklistItemCount", 0) > 0,
        })

    return pd.DataFrame(rows)


def _priority_label(p: int) -> str:
    return {0: "Urgent", 1: "Important", 2: "Medium", 3: "Low"}.get(p // 2, str(p))


# ─── Archive + write ──────────────────────────────────────────────────────────
def archive_existing(name: str, ts: str) -> None:
    """Move existing raw file → archive with timestamp suffix."""
    existing = RAW_DIR / f"{name}_Claims.xlsx"
    if existing.exists():
        dest = ARCH_DIR / f"{name}_Claims_{ts}.xlsx"
        shutil.move(str(existing), str(dest))
        print(f"    Archived → archive/{dest.name}")


# ─── Main ─────────────────────────────────────────────────────────────────────
def main() -> None:
    if CLIENT_ID == "YOUR_APP_CLIENT_ID":
        sys.exit(
            "CLIENT_ID not set.\n"
            "Complete the Azure AD app registration (see SETUP_AUTH.md), then\n"
            "paste the Application (client) ID into this script."
        )

    RAW_DIR.mkdir(parents=True, exist_ok=True)
    ARCH_DIR.mkdir(parents=True, exist_ok=True)

    print("Authenticating with Microsoft Graph...")
    token = get_token()
    print("  Auth OK\n")

    ts = datetime.now().strftime("%Y%m%d_%H%M")

    for name, plan_id in PLANS.items():
        print(f"[{name} Claims]")
        archive_existing(name, ts)
        df = extract_plan(token, plan_id)
        out = RAW_DIR / f"{name}_Claims.xlsx"
        df.to_excel(out, index=False)
        print(f"    {len(df)} tasks → raw/{out.name}\n")

    print("Done.")


if __name__ == "__main__":
    main()
