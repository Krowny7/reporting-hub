from __future__ import annotations

from pathlib import Path

APP_TITLE = "Reporting Hub"
DEFAULT_PILOT_MACRO = "Run_MonthEnd_Update"

# Report frequencies shown in the Update page.
# UI uses the label, settings store the lower-case key (e.g. "monthly").
REPORT_TYPE_OPTIONS = ["Weekly", "Monthly", "Quarterly", "Semiannual"]
DEFAULT_REPORT_TYPE = "monthly"

# settings.json is resolved from current working directory (portable when packaged)
SETTINGS_PATH = Path.cwd() / "settings.json"
