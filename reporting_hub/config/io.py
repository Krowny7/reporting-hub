from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

from .models import MacroDefinition, Settings


def _parse_macros(raw: Any) -> Dict[str, MacroDefinition]:
    if not isinstance(raw, dict):
        return {}

    out: Dict[str, MacroDefinition] = {}
    for macro_id, item in raw.items():
        if not isinstance(item, dict):
            continue

        label = str(item.get("label", macro_id))
        workbook_path = str(item.get("workbook_path", ""))
        macro = str(item.get("macro", ""))
        args = str(item.get("args", ""))

        if macro.strip():
            out[str(macro_id)] = MacroDefinition(
                label=label,
                workbook_path=workbook_path,
                macro=macro,
                args=args,
            )

    return out


def load_settings(path: Path) -> Settings:
    """Load settings from JSON (missing file -> defaults)."""
    if not path.exists():
        return Settings()

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return Settings()

    if not isinstance(data, dict):
        return Settings()

    s = Settings()
    s.appearance = str(data.get("appearance", s.appearance))
    s.excel_mode = str(data.get("excel_mode", s.excel_mode))

    # Selected report frequency (stored as a lower-case key)
    s.report_type = str(data.get("report_type", s.report_type)).strip().lower() or s.report_type

    # Backward-compatible keys
    s.pilot_path = str(data.get("pilot_path", s.pilot_path))
    s.pilot_macro = str(data.get("pilot_macro", s.pilot_macro))
    s.pilot_args = str(data.get("pilot_args", s.pilot_args))

    s.macros = _parse_macros(data.get("macros"))
    return s


def _settings_to_dict(settings: Settings) -> Dict[str, Any]:
    return {
        "appearance": settings.appearance,
        "excel_mode": settings.excel_mode,
        "report_type": settings.report_type,
        "pilot_path": settings.pilot_path,
        "pilot_macro": settings.pilot_macro,
        "pilot_args": settings.pilot_args,
        "macros": {
            macro_id: {
                "label": m.label,
                "workbook_path": m.workbook_path,
                "macro": m.macro,
                "args": m.args,
            }
            for macro_id, m in settings.macros.items()
        },
    }


def save_settings(path: Path, settings: Settings) -> None:
    """Save settings as JSON (best effort)."""
    try:
        payload = _settings_to_dict(settings)
        path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    except Exception:
        # Best effort: do not crash UI
        return
