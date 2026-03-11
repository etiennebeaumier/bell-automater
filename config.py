"""Application configuration manager with JSON persistence."""

import json
import os
import pathlib
from datetime import datetime

CONFIG_DIR = pathlib.Path.home() / ".bcecn_pricing"
CONFIG_FILE = CONFIG_DIR / "config.json"

SHAREPOINT_AUTOMATION_RELATIVE = (
    pathlib.Path("Bell Canada")
    / "MTRLPQDC304 - Capit_Market"
    / "Automation"
)
DEFAULT_MASTER_FILENAME = "Master File - EB.xlsm"
DEFAULT_PDF_SUBDIR = "PDF"


def _sharepoint_root_candidates() -> list[pathlib.Path]:
    """Return likely local roots for SharePoint/OneDrive-synced folders."""
    candidates = [pathlib.Path.home()]
    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        candidates.append(pathlib.Path(userprofile))

    deduped: list[pathlib.Path] = []
    seen = set()
    for candidate in candidates:
        key = str(candidate).lower()
        if key not in seen:
            deduped.append(candidate)
            seen.add(key)
    return deduped


def get_default_automation_dir() -> str:
    """Resolve the best local SharePoint automation folder for this user profile."""
    for root in _sharepoint_root_candidates():
        candidate = root / SHAREPOINT_AUTOMATION_RELATIVE
        if candidate.exists():
            return str(candidate)
    return str(_sharepoint_root_candidates()[0] / SHAREPOINT_AUTOMATION_RELATIVE)


def get_default_pdf_dir() -> str:
    """Return the default local folder used for PDF intake."""
    return str(pathlib.Path(get_default_automation_dir()) / DEFAULT_PDF_SUBDIR)


def get_default_master_file() -> str:
    """Return the default master workbook path."""
    return str(pathlib.Path(get_default_automation_dir()) / DEFAULT_MASTER_FILENAME)

DEFAULTS = {
    "master_file": get_default_master_file(),
    "pdf_source_dir": get_default_pdf_dir(),
    "appearance_mode": "dark",
    "dry_run": False,
    "avg_start_year": datetime.now().year,
    "avg_end_year": datetime.now().year,
}


class AppConfig:
    """JSON-backed configuration storage for desktop app preferences."""

    def __init__(self):
        """Initialize config with defaults, then migrate/load persisted values."""
        self._data = dict(DEFAULTS)
        self._migrate_env()
        self.load()

    def _migrate_env(self):
        """On first run, seed config from .env if it exists and no config file yet."""
        if CONFIG_FILE.exists():
            return
        env_path = pathlib.Path(__file__).parent / ".env"
        if not env_path.exists():
            return
        env_vals = {}
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, value = line.split("=", 1)
                key, value = key.strip(), value.strip()
                if (value.startswith('"') and value.endswith('"')) or (
                    value.startswith("'") and value.endswith("'")
                ):
                    value = value[1:-1]
                if key:
                    env_vals[key] = value

        mapping = {
            "MASTER_FILE": "master_file",
            "PDF_SOURCE_DIR": "pdf_source_dir",
        }
        for env_key, cfg_key in mapping.items():
            if env_key in env_vals and env_vals[env_key]:
                self._data[cfg_key] = env_vals[env_key]
        self.save()

    def load(self):
        """Load persisted config values from disk when available."""
        if CONFIG_FILE.exists():
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    stored = json.load(f)
                for key in DEFAULTS:
                    if key in stored:
                        self._data[key] = stored[key]
            except (json.JSONDecodeError, OSError):
                pass

    def save(self):
        """Persist current config values to disk."""
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self._data, f, indent=2)

    def get(self, key, default=None):
        """Read a config value with optional fallback."""
        return self._data.get(key, default)

    def set(self, key, value):
        """Set a config value in memory."""
        self._data[key] = value

    def __getitem__(self, key):
        """Dictionary-style key access."""
        return self._data[key]

    def __setitem__(self, key, value):
        """Dictionary-style key assignment."""
        self._data[key] = value
