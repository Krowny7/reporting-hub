from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict


@dataclass
class MacroDefinition:
    """A runnable macro definition."""

    label: str
    workbook_path: str
    macro: str
    args: str = ""  # semicolon-separated


@dataclass
class Settings:
    """Application settings stored in settings.json."""

    appearance: str = "Dark"
    excel_mode: str = "minimized"  # minimized | hidden | visible

    # Pilot (backward-compatible with your current UI)
    pilot_path: str = ""
    pilot_macro: str = "Run_MonthEnd_Update"
    pilot_args: str = ""

    # Optional registry for multiple macros
    macros: Dict[str, MacroDefinition] = field(default_factory=dict)
