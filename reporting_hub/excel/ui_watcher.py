from __future__ import annotations

import threading
import time

try:
    import win32con
    import win32gui
    import win32process
except Exception:  # pragma: no cover
    win32con = None
    win32gui = None
    win32process = None


class ExcelUIWatcher:
    """
    Goal: keep Excel main window minimized/hidden while letting MsgBox/UserForms appear.

    Key stability rule (anti-flicker):
    - NEVER "bring to front" a dialog repeatedly.
    - Enforce main window state only:
        * periodically when no dialog exists,
        * once when a dialog first appears (to hide the workbook),
        * then pause enforcement while the dialog is open.
    """

    def __init__(self, excel_pid: int, main_mode: str = "minimized"):
        self.excel_pid = excel_pid
        self._stop = threading.Event()
        self._thread = None
        self._main_mode = (main_mode or "minimized").strip().lower()

        # Anti-flicker state
        self._had_dialogs = False
        self._seen_dialogs: set[int] = set()
        self._last_main_enforce = 0.0

        # Tunables
        self._poll_s = 0.45              # lower frequency = less redraw/flicker
        self._main_enforce_period_s = 1.5  # only enforce main window every Xs (when no dialog)

    def set_main_mode(self, mode: str) -> None:
        m = (mode or "").strip().lower()
        if m not in ("minimized", "hidden", "visible"):
            m = "minimized"
        self._main_mode = m

    def start(self) -> None:
        if not win32gui or not win32process or not win32con:
            return
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def stop(self) -> None:
        self._stop.set()

    def _iter_excel_windows(self):
        hwnds = []

        def enum_cb(hwnd, _):
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                if pid != self.excel_pid:
                    return True
                if not win32gui.IsWindow(hwnd):
                    return True
                hwnds.append(hwnd)
            except Exception:
                pass
            return True

        win32gui.EnumWindows(enum_cb, None)
        return hwnds

    def _class_name(self, hwnd) -> str:
        try:
            return win32gui.GetClassName(hwnd)
        except Exception:
            return ""

    def _is_main_excel_window(self, hwnd) -> bool:
        cls = self._class_name(hwnd)
        return cls in ("XLMAIN", "EXCEL7")

    def _is_dialog_or_userform(self, hwnd) -> bool:
        cls = self._class_name(hwnd)
        if cls in ("XLMAIN", "EXCEL7"):
            return False
        if cls == "#32770":            # standard dialog
            return True
        if cls.startswith("Thunder"):  # VBA UserForms
            return True
        return False

    def _enforce_main_window_state(self, hwnd) -> None:
        """Keep the workbook UI out of sight, but do it gently (only if needed)."""
        try:
            if self._main_mode == "visible":
                return

            if self._main_mode == "hidden":
                # Hide only if currently visible
                if win32gui.IsWindowVisible(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
                return

            # minimized
            if not win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        except Exception:
            pass

    def _bring_dialog_once(self, hwnd) -> None:
        """
        Bring dialog/userform to the front ONCE (no repeated topmost toggles).
        This avoids stealing focus while user clicks OK.
        """
        try:
            # Only restore/show if actually minimized/hidden
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            elif not win32gui.IsWindowVisible(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_SHOW)

            # Try normal foreground
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                pass

            # If still not foreground, do ONE topmost punch-through (still single-shot)
            try:
                fg = win32gui.GetForegroundWindow()
            except Exception:
                fg = None

            if fg != hwnd:
                try:
                    win32gui.SetWindowPos(
                        hwnd,
                        win32con.HWND_TOPMOST,
                        0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW,
                    )
                    win32gui.SetWindowPos(
                        hwnd,
                        win32con.HWND_NOTOPMOST,
                        0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_SHOWWINDOW,
                    )
                    win32gui.SetForegroundWindow(hwnd)
                except Exception:
                    pass
        except Exception:
            pass

    def _run(self) -> None:
        while not self._stop.is_set():
            try:
                windows = self._iter_excel_windows()
                main_hwnds = [h for h in windows if self._is_main_excel_window(h)]
                dialog_hwnds = [h for h in windows if self._is_dialog_or_userform(h)]

                dialogs_present = len(dialog_hwnds) > 0

                # If a dialog just appeared: enforce main state ONCE (hide workbook),
                # then stop touching main window while dialog exists (prevents flicker).
                if dialogs_present and not self._had_dialogs:
                    for mh in main_hwnds:
                        self._enforce_main_window_state(mh)
                    self._had_dialogs = True
                    # reset so we can "single-shot" dialogs
                    self._seen_dialogs.clear()

                # Handle dialogs: bring each dialog/userform ONLY ONCE per appearance.
                if dialogs_present:
                    for dh in dialog_hwnds:
                        dhi = int(dh)
                        if dhi not in self._seen_dialogs:
                            self._seen_dialogs.add(dhi)
                            self._bring_dialog_once(dh)
                else:
                    # No dialogs: periodically enforce main window state (gentle)
                    now = time.monotonic()
                    if (now - self._last_main_enforce) >= self._main_enforce_period_s:
                        for mh in main_hwnds:
                            self._enforce_main_window_state(mh)
                        self._last_main_enforce = now

                    # reset dialog state
                    self._had_dialogs = False
                    self._seen_dialogs.clear()

            except Exception:
                pass

            time.sleep(self._poll_s)
