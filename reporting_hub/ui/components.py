# reporting_hub/ui/components.py
"""
Compat layer: allows older imports like `from reporting_hub.ui.components import ...`
while the real implementation lives in `reporting_hub.gui`.
"""

from __future__ import annotations

# Re-export style tokens + helpers
from ..gui.style import (  # noqa: F401
    C,
    BG_APP,
    SIDEBAR,
    CARD,
    FIELD,
    BORDER,
    TEXT,
    MUTED,
    BTN_MAIN_BG,
    BTN_MAIN_TEXT,
    BTN_MAIN_HOVER,
    BTN_GHOST_BG,
    BTN_GHOST_HOVER,
    apply_app_style,
    font,
)

# Re-export UI widgets
from ..gui.widgets import Card, ToastHost, btn_primary, btn_ghost  # noqa: F401

__all__ = [
    # style
    "C",
    "BG_APP",
    "SIDEBAR",
    "CARD",
    "FIELD",
    "BORDER",
    "TEXT",
    "MUTED",
    "BTN_MAIN_BG",
    "BTN_MAIN_TEXT",
    "BTN_MAIN_HOVER",
    "BTN_GHOST_BG",
    "BTN_GHOST_HOVER",
    "apply_app_style",
    "font",
    # widgets
    "Card",
    "ToastHost",
    "btn_primary",
    "btn_ghost",
]
