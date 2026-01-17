from __future__ import annotations

import argparse
import sys

from pathlib import Path

from .app import App
from .config.constants import SETTINGS_PATH
from .config.io import load_settings
from .services.macro_runner import MacroRunner, RunRequest


def _parse_args(argv: list[str]) -> argparse.Namespace:
    p = argparse.ArgumentParser(prog="reporting_hub")
    p.add_argument("--headless", action="store_true", help="Run without GUI")
    p.add_argument("--list", action="store_true", help="List macro ids from settings.json")
    p.add_argument("--macro", dest="macro_id", default="", help="Macro id from settings.json (macros.<id>)")
    p.add_argument("--pilot", dest="pilot_path", default="", help="Workbook path (overrides settings)")
    p.add_argument("--macro-name", dest="macro_name", default="", help="Macro name (overrides settings)")
    p.add_argument("--args", dest="args", default="", help="Args separated by ';'")
    p.add_argument("--excel-mode", dest="excel_mode", default="", help="minimized|hidden|visible")
    p.add_argument("--quit-excel", action="store_true", help="Quit Excel after running (headless)")
    return p.parse_args(argv)


def _log_to_stdout(msg: str) -> None:
    print(msg, flush=True)


def main(argv: list[str] | None = None) -> int:
    ns = _parse_args(list(argv) if argv is not None else sys.argv[1:])

    settings = load_settings(SETTINGS_PATH)

    if ns.list:
        if not settings.macros:
            print("No macros declared in settings.json under 'macros'.")
            return 0
        for macro_id, m in settings.macros.items():
            print(f"{macro_id}: {m.label} -> {m.macro}")
        return 0

    if ns.headless:
        # Resolve request from CLI overrides / settings
        if ns.macro_id:
            if ns.macro_id not in settings.macros:
                print(f"Unknown macro id: {ns.macro_id}")
                return 2
            m = settings.macros[ns.macro_id]
            workbook_path = (ns.pilot_path or m.workbook_path or settings.pilot_path).strip()
            macro_name = (ns.macro_name or m.macro or settings.pilot_macro).strip()
            raw_args = ns.args if ns.args else (m.args or settings.pilot_args)
        else:
            workbook_path = (ns.pilot_path or settings.pilot_path).strip()
            macro_name = (ns.macro_name or settings.pilot_macro).strip()
            raw_args = ns.args if ns.args else settings.pilot_args

        args = [a.strip() for a in str(raw_args).split(";") if a.strip()]
        excel_mode = (ns.excel_mode or settings.excel_mode or "minimized").strip().lower()

        if not workbook_path:
            print("Missing workbook path. Use --pilot or set 'pilot_path' in settings.json.")
            return 2
        if not macro_name:
            print("Missing macro name. Use --macro-name or set 'pilot_macro' in settings.json.")
            return 2

        runner = MacroRunner(_log_to_stdout)
        runner.run(
            RunRequest(workbook_path=workbook_path, macro_name=macro_name, args=args, excel_mode=excel_mode),
            quit_excel_when_done=bool(ns.quit_excel),
        )
        return 0

    # GUI
    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
