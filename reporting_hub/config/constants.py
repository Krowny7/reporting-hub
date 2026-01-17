from __future__ import annotations

from pathlib import Path

APP_TITLE = "Reporting Hub"
DEFAULT_PILOT_MACRO = "Run_MonthEnd_Update"

# settings.json is resolved from current working directory (portable when packaged)
SETTINGS_PATH = Path.cwd() / "settings.json"
