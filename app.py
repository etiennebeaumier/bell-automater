#!/usr/bin/env python3
"""BCECN Pricing Tool - Desktop application entry point."""

import sys
import os

# Ensure the project root is on sys.path (needed when running from PyInstaller bundle)
if getattr(sys, "frozen", False):
    os.chdir(os.path.dirname(sys.executable))
    sys.path.insert(0, os.path.dirname(sys.executable))
else:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import AppConfig
from ui.app_window import AppWindow


def main():
    config = AppConfig()
    app = AppWindow(config)
    app.mainloop()


if __name__ == "__main__":
    main()
