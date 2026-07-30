"""
download_planner_browser.py — Export Claims Planner boards via Chrome automation
=================================================================================
Uses Playwright (bundled Chromium) with saved M365 auth state to click
"Export as Excel" on the Ford and Chrysler Planner boards.

PRIMARY automation path. Supersedes download_planner.py (MSAL fallback).

FIRST-TIME SETUP (one-time, ~2 minutes):
    python scripts/download_planner_browser.py --setup
    A browser window opens → sign in to M365 → close the window.
    Auth state is saved to data/.auth.json and reused on every future run.

    Re-run --setup if you see login prompts during a scheduled run (tokens
    typically last 90 days).

NORMAL USAGE:
    python scripts/download_planner_browser.py          # headless (default)
    python scripts/download_planner_browser.py --headed # visible browser
    python scripts/download_planner_browser.py --plan Ford  # single plan

OUTPUT:
    data/raw/Ford_Claims.xlsx
    data/raw/Chrysler_Claims.xlsx
    data/archive/Ford_Claims_<timestamp>.xlsx  (previous file archived)
"""

from __future__ import annotations

import argparse
import shutil
import time
import sys
from datetime import datetime
from pathlib import Path

# ─── Configuration ─────────────────────────────────────────────────────────────

PLANS = {
    "Ford":     "jRdOeWc20kuFagvcx_OkrmUABPJd",
    "Chrysler": "hYGA2uMjlE-9NFjXfoUFeWUAGOP3",
}

PLANNER_BASE = "https://planner.cloud.microsoft/webui/plan"
M365_LOGIN   = "https://planner.cloud.microsoft"

HERE      = Path(__file__).resolve().parent.parent   # project root (2. Claims/)
RAW_DIR   = HERE / "data" / "raw"
ARCH_DIR  = HERE / "data" / "archive"
AUTH_FILE = HERE / "data" / ".auth.json"

PAGE_LOAD_TIMEOUT = 45_000   # ms
EXPORT_TIMEOUT    = 90_000   # ms


# ─── Helpers ───────────────────────────────────────────────────────────────────

def archive_existing(label: str, ts: str) -> None:
    """Move current raw file to archive with timestamp."""
    for ext in [".xlsx", ".xls"]:
        f = RAW_DIR / f"{label}_Claims{ext}"
        if f.exists():
            dest = ARCH_DIR / f"{label}_Claims_{ts}{ext}"
            shutil.move(str(f), str(dest))
            print(f"    Archived → archive/{dest.name}")
            return


def export_plan(page, plan_id: str, label: str) -> None:
    """Navigate to a Planner board and trigger Export as Excel."""
    url = f"{PLANNER_BASE}/{plan_id}"
    print(f"  Loading {label} Claims Planner...")

    page.goto(url, wait_until="domcontentloaded", timeout=PAGE_LOAD_TIMEOUT)

    # Bail out if we ended up on a login page
    if "login" in page.url.lower() or "microsoftonline" in page.url.lower():
        raise RuntimeError(
            "Redirected to login page — auth tokens expired.\n"
            "Run:  python scripts/download_planner_browser.py --setup"
        )

    # Wait for the board to render — wait for the plan title in the header
    page.wait_for_selector("text=Claims", timeout=PAGE_LOAD_TIMEOUT)

    # Let the board settle — wait for at least one task card to appear
    page.wait_for_selector("text=Add task", timeout=PAGE_LOAD_TIMEOUT)
    time.sleep(3)

    _click_more_menu(page)

    page.wait_for_selector("text=Export as Excel", timeout=10_000)

    with page.expect_download(timeout=EXPORT_TIMEOUT) as dl_info:
        page.click("text=Export as Excel")

    download = dl_info.value
    dest = RAW_DIR / f"{label}_Claims.xlsx"
    download.save_as(str(dest))

    size_kb = dest.stat().st_size // 1024
    print(f"    Saved → raw/{dest.name}  ({size_kb} KB)")


def _click_more_menu(page) -> None:
    """Click the ⋯ plan menu button in the header using a real mouse click."""
    # Get the button's screen coordinates — JS click doesn't fire React events.
    # Restrict to header band (y: 80–180px) with MoreHorizontal icon.
    coords = page.evaluate("""
        () => {
            // The plan ··· button has aria-label ending with "– Plan options"
            // e.g. "Ford Claims Planner – Plan options"
            // This is stable across Ford/Chrysler and different from task buttons.
            const btn = [...document.querySelectorAll('button')].find(b =>
                (b.getAttribute('aria-label') || '').endsWith('– Plan options') ||
                (b.getAttribute('aria-label') || '').endsWith('- Plan options') ||
                (b.getAttribute('aria-label') || '').endsWith('Plan options')
            );
            if (!btn) return null;
            const r = btn.getBoundingClientRect();
            return {
                x: r.x + r.width / 2,
                y: r.y + r.height / 2,
                label: btn.getAttribute('aria-label') || '(no label)'
            };
        }
    """)

    if not coords:
        raise RuntimeError(
            "Could not find the ⋯ plan menu button in the header (y: 80–180px).\n"
            "Run with --headed to inspect; update y-range in _click_more_menu() if needed."
        )

    print(f"    Clicking plan menu button at ({coords['x']:.0f}, {coords['y']:.0f}) "
          f"label={coords['label']!r}")
    page.mouse.click(coords["x"], coords["y"])


# ─── Setup mode ────────────────────────────────────────────────────────────────

def run_setup(p) -> None:
    """
    Headed browser → user logs in → saves auth state.
    Run once, then normal runs reuse the saved session.
    """
    print("SETUP MODE — Sign in to M365")
    print("="*50)
    print("1. A browser window will open.")
    print("2. Sign in with your Microsoft account.")
    print("3. Navigate to any Planner page to confirm access.")
    print("4. Close the browser window when done.")
    print("="*50)
    input("Press Enter to open the browser...")

    browser = p.chromium.launch(headless=False, channel=None)
    context = browser.new_context()
    page = context.new_page()
    page.goto(M365_LOGIN)

    print("\nBrowser open. Sign in, then close the browser window when ready.")
    print("(waiting for you to close the browser...)")

    try:
        # Wait until the user closes the browser
        page.wait_for_event("close", timeout=0)
    except Exception:
        pass

    # Save auth state before closing
    AUTH_FILE.parent.mkdir(parents=True, exist_ok=True)
    context.storage_state(path=str(AUTH_FILE))
    print(f"\nAuth state saved → {AUTH_FILE.relative_to(HERE)}")
    print("You can now run the script without --setup.")
    browser.close()


# ─── Main ───────────────────────────────────────────────────────────────────────

def main() -> None:
    ap = argparse.ArgumentParser(description="Download Claims Planner boards via Playwright")
    ap.add_argument("--setup", action="store_true",
                    help="One-time: open browser, sign in to M365, save auth state")
    ap.add_argument("--headed", action="store_true",
                    help="Show the browser window (default: headless)")
    ap.add_argument("--plan", choices=list(PLANS.keys()),
                    help="Export only one plan (default: both)")
    args = ap.parse_args()

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        sys.exit(
            "Playwright not installed.\n"
            "Run:  pip install playwright\n"
            "      playwright install chromium"
        )

    with sync_playwright() as p:
        if args.setup:
            run_setup(p)
            return

        # Normal run — require saved auth state
        if not AUTH_FILE.exists():
            sys.exit(
                "No auth state found. Run setup first:\n"
                "  python scripts/download_planner_browser.py --setup"
            )

        RAW_DIR.mkdir(parents=True, exist_ok=True)
        ARCH_DIR.mkdir(parents=True, exist_ok=True)

        ts = datetime.now().strftime("%Y%m%d_%H%M")
        plans = {args.plan: PLANS[args.plan]} if args.plan else PLANS

        print("Starting Planner export...")

        browser = p.chromium.launch(headless=not args.headed)
        context = browser.new_context(
            storage_state=str(AUTH_FILE),
            accept_downloads=True,
            viewport={"width": 1600, "height": 900},
        )
        page = context.new_page()
        page.set_default_timeout(PAGE_LOAD_TIMEOUT)

        for label, plan_id in plans.items():
            print(f"\n[{label} Claims]")
            try:
                archive_existing(label, ts)
                export_plan(page, plan_id, label)
            except Exception as e:
                # Save screenshot to help diagnose
                shot = HERE / "data" / f"{label}_error_{ts}.png"
                try:
                    page.screenshot(path=str(shot), full_page=False)
                    print(f"    Screenshot → {shot.relative_to(HERE)}")
                except Exception:
                    pass
                print(f"    ERROR: {e}")
                print(f"    Skipping {label} — manual download required.")

        browser.close()

    print("\nDone.")
    print("Files in data/raw/:")
    for f in sorted(RAW_DIR.glob("*_Claims.xlsx")):
        print(f"  {f.name}  ({f.stat().st_size // 1024} KB)")


if __name__ == "__main__":
    main()
