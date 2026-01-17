from __future__ import annotations

from pathlib import Path

APP_TITLE = "Reporting Hub"

# Default macro if nothing else is set (legacy / monthly)
DEFAULT_PILOT_MACRO = "Run_MonthEnd_Update"

# Frequencies shown in UI
REPORT_TYPE_OPTIONS = ["Weekly", "Monthly", "Quarterly", "Semiannual"]
DEFAULT_REPORT_TYPE = "monthly"

# Default macro per frequency (used to auto-fill profiles)
REPORT_TYPE_DEFAULT_MACROS = {
    "weekly": "Run_Weekly_Update",
    "monthly": "Run_MonthEnd_Update",
    "quarterly": "Run_Quarterly_Update",
    "semiannual": "Run_Semiannual_Update",
}

SETTINGS_PATH = Path.cwd() / "settings.json"
