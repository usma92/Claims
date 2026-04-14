"""
src/paths.py
------------
Single source of truth for all filesystem paths in the project.
All notebooks and scripts resolve paths through get_paths() — no hardcoded
paths anywhere else in the codebase.
"""
from pathlib import Path


def get_paths() -> dict:
    """Return absolute paths for every data layer and output directory."""
    root = Path(__file__).resolve().parents[1]
    return {
        "root":      root,
        "config":    root / "config",
        "raw":       root / "data" / "inputs",
        "interim":   root / "data" / "interim",
        "processed": root / "data" / "processed",
        "exports":   root / "data" / "exports",
        "charts":    root / "figs_claims",
        "reports":   root / "reports",
    }
