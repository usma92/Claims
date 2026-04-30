# src/paths.py -- Centralized path resolver for Claims Aging Pipeline.
from pathlib import Path


def get_paths() -> dict:
    root = Path(__file__).resolve().parents[1]
    return {
        "root":       root,
        "config":     root / "config",
        "raw":        root / "data" / "raw",
        "interim":    root / "data" / "interim",
        "processed":  root / "data" / "processed",
        "exports":    root / "data" / "exports",
        "charts":     root / "charts",
        "dashboards": root / "dashboards",
        "reports":    root / "reports",
    }
