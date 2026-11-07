import sys
from pathlib import Path


def app_dir() -> Path:
    """Writable directory for logs/sessions/db - lives beside the exe when frozen."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resource_dir() -> Path:
    """Read-only bundled resources (templates, static, sample data)."""
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent
