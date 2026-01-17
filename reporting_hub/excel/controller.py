from __future__ import annotations

import os
import time
from typing import Callable, Optional

try:
    import pythoncom
    import win32com.client
    import win32gui
    import win32con
    import win32process
except Exception:  # pragma: no cover
    pythoncom = None
    win32com = None
    win32gui = None
    win32con = None
    win32process = None

from .ui_watcher import ExcelUIWatcher


Logger = Callable[[str], None]


class ExcelController:
    """Thin wrapper around Excel COM.

    - creates a dedicated Excel instance
    - opens/activates workbooks
    - runs macros
    - optionally keeps dialogs visible via ExcelUIWatcher
    """

    def __init__(self, logger: Logger):
        self.logger = logger
        self.excel = None
        self.excel_pid: Optional[int] = None
        self.ui_watcher: Optional[ExcelUIWatcher] = None
        self.mode = "minimized"  # minimized | hidden | visible

    def _log(self, msg: str) -> None:
        try:
            self.logger(msg)
        except Exception:
            pass

    def _ensure_excel(self) -> None:
        if self.excel is None:
            raise RuntimeError("Excel n'est pas lancé.")

    def launch_new_instance(self) -> None:
        if not pythoncom or not win32com:
            raise RuntimeError("pywin32 est requis (Windows uniquement).")

        # NOTE:
        # Excel COM objects MUST be created and used on the same thread.
        # We initialize COM once in the dedicated Excel worker thread.
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self._log("Excel: instance dédiée lancée.")

        # Best effort to reduce prompts
        for attr, value in (("DisplayAlerts", False), ("AskToUpdateLinks", False)):
            try:
                setattr(self.excel, attr, value)
            except Exception:
                pass

        # Setup watcher
        try:
            hwnd = self.excel.Hwnd
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            self.excel_pid = pid
            # Keep the *main* Excel window discreet, but still allow dialogs
            # (MsgBox/UserForms) to surface via the watcher.
            self.ui_watcher = ExcelUIWatcher(pid, main_mode=self.mode)
            self.ui_watcher.start()
            self._log(f"Excel: watcher UI actif (PID={pid}).")
        except Exception:
            self._log("Excel: watcher UI non initialisé (pas bloquant).")

        self.set_excel_mode(self.mode)

    def quit_excel(self) -> None:
        self._ensure_excel()
        try:
            if self.ui_watcher:
                self.ui_watcher.stop()
            try:
                self.excel.DisplayAlerts = False
            except Exception:
                pass
            self.excel.Quit()
            self._log("Excel: fermé.")
        finally:
            self.excel = None
            self.excel_pid = None
            self.ui_watcher = None

    def set_excel_mode(self, mode: str) -> None:
        self._ensure_excel()
        mode = (mode or "").strip().lower()
        if mode not in ("minimized", "hidden", "visible"):
            mode = "minimized"
        self.mode = mode

        # Tell the watcher what we want for the MAIN Excel window.
        # (Dialogs/UserForms are still allowed to pop to the front.)
        try:
            if self.ui_watcher:
                self.ui_watcher.set_main_mode(mode)
        except Exception:
            pass

        hwnd = None
        try:
            hwnd = self.excel.Hwnd
        except Exception:
            pass

        try:
            if mode == "hidden":
                # Keep Excel "Visible" at the COM level so dialogs/UserForms
                # can still surface, but hide the main window.
                self.excel.Visible = True
                if hwnd and win32gui:
                    win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
            elif mode == "minimized":
                self.excel.Visible = True
                if hwnd and win32gui:
                    win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            else:
                self.excel.Visible = True
                if hwnd and win32gui:
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        except Exception:
            pass

    def show_excel_for_seconds(self, seconds: int = 10) -> None:
        """Show Excel temporarily.

        IMPORTANT:
        Avoid spawning a new thread here: COM calls must remain on the same
        thread as the Excel instance.

        This method therefore only shows Excel. Restoring the previous mode is
        handled by the Excel worker (which can schedule a delayed action).
        """
        self._ensure_excel()
        self.set_excel_mode("visible")

        # Small sleep to make the window state change visible in edge cases.
        time.sleep(0.05)

    def open_or_activate_by_path(self, path: str) -> str:
        self._ensure_excel()
        path = os.path.abspath(path)
        if not os.path.exists(path):
            raise RuntimeError(f"Classeur introuvable: {path}")

        # Try already open
        try:
            wb = self.excel.Workbooks(path)
            try:
                wb.Activate()
            except Exception:
                pass
            return wb.Name
        except Exception:
            pass

        # Otherwise open
        try:
            try:
                wb = self.excel.Workbooks.Open(path, UpdateLinks=0)
            except Exception:
                wb = self.excel.Workbooks.Open(path)
            return wb.Name
        except Exception as e:
            raise RuntimeError(f"Erreur ouverture pilote: {e}")

    def run_macro(self, wb_name: str, macro_name: str, *args) -> None:
        self._ensure_excel()
        macro_name = (macro_name or "").strip()
        if not macro_name:
            raise RuntimeError("Nom de macro vide.")

        attempts = []
        if "!" in macro_name:
            attempts.append(macro_name)
        else:
            attempts.append(macro_name)
            attempts.append(f"{wb_name}!{macro_name}")

        last_err = None
        for m in attempts:
            try:
                self.excel.Application.Run(m, *args)
                self._log(f"Macro OK: {m}")
                return
            except Exception as e:
                last_err = e

        raise RuntimeError(f"Macro KO. Dernière erreur: {last_err}")
