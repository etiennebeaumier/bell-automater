"""Application configuration manager with JSON persistence."""

import json
import os
import pathlib

CONFIG_DIR = pathlib.Path.home() / ".bcecn_pricing"
CONFIG_FILE = CONFIG_DIR / "config.json"

DEFAULTS = {
    "master_file": "",
    "outlook_email": "",
    "outlook_server": "outlook.office365.com",
    "outlook_days": 7,
    "bcecn_sender": "",
    "appearance_mode": "dark",
    "dry_run": False,
}


class AppConfig:
    def __init__(self):
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
            "OUTLOOK_EMAIL": "outlook_email",
            "OUTLOOK_SERVER": "outlook_server",
            "OUTLOOK_DAYS": "outlook_days",
            "BCECN_SENDER": "bcecn_sender",
        }
        for env_key, cfg_key in mapping.items():
            if env_key in env_vals and env_vals[env_key]:
                val = env_vals[env_key]
                if cfg_key == "outlook_days":
                    try:
                        val = int(val)
                    except ValueError:
                        continue
                self._data[cfg_key] = val
        self.save()

    def load(self):
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
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self._data, f, indent=2)

    def get(self, key, default=None):
        return self._data.get(key, default)

    def set(self, key, value):
        self._data[key] = value

    def __getitem__(self, key):
        return self._data[key]

    def __setitem__(self, key, value):
        self._data[key] = value
