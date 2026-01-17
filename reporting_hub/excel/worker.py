from __future__ import annotations

import queue
import threading
import time
import traceback
from dataclasses import dataclass
from typing import Any, Callable, Optional

try:
    import pythoncom
except Exception:  # pragma: no cover
    pythoncom = None

from .controller import ExcelController


UIFn = Callable[..., None]


@dataclass
class _Task:
    action: str
    args: tuple
    kwargs: dict
    on_ok: Optional[Callable[[Any], None]] = None
    on_err: Optional[Callable[[BaseException], None]] = None


class ExcelWorker:
    """Runs all Excel COM operations on ONE dedicated thread.

    Why:
    - Excel COM objects are thread-affine (create + use in the same thread)
    - long-running macros (30-60min) must NOT freeze the UI

    The UI thread communicates via a queue of tasks.
    """

    def __init__(self, ui_root, ui_log: UIFn, ui_toast: UIFn):
        self._ui_root = ui_root
        self._ui_log = ui_log
        self._ui_toast = ui_toast

        self._q: "queue.Queue[_Task]" = queue.Queue()
        self._stop = threading.Event()

        self._thread = threading.Thread(target=self._run, name="ExcelWorker", daemon=True)
        self._thread.start()

    # ------------------------------
    # Public API (called from UI)
    # ------------------------------
    def submit(self, action: str, *args, on_ok=None, on_err=None, **kwargs) -> None:
        self._q.put(_Task(action=action, args=args, kwargs=kwargs, on_ok=on_ok, on_err=on_err))

    def start(self) -> None:
        """Compatibility no-op.

        The worker thread is started automatically in __init__.
        """
        return None

    def stop(self) -> None:
        """Stops the worker thread (best effort)."""
        self._stop.set()
        self._q.put(_Task(action="__stop__", args=(), kwargs={}))

    # ------------------------------
    # Internal
    # ------------------------------
    def _ui(self, fn: Callable, *args, **kwargs) -> None:
        try:
            self._ui_root.after(0, lambda: fn(*args, **kwargs))
        except Exception:
            # If UI is gone, ignore.
            pass

    def _run(self) -> None:
        if pythoncom is None:  # pragma: no cover
            controller: Optional[ExcelController] = None
        else:
            pythoncom.CoInitialize()
            controller = ExcelController(logger=lambda m: self._ui(self._ui_log, m))

        while not self._stop.is_set():
            try:
                task = self._q.get(timeout=0.25)
            except queue.Empty:
                continue

            if task.action == "__stop__":
                break

            try:
                if controller is None:
                    raise RuntimeError("pywin32 est requis (Windows uniquement).")

                result = self._dispatch(controller, task)
                if task.on_ok:
                    self._ui(task.on_ok, result)
            except BaseException:
                if task.on_err:
                    tb = traceback.format_exc()
                    self._ui(task.on_err, RuntimeError(tb))
                else:
                    self._ui(self._ui_toast, "Excel error (see logs).")

        # Best effort cleanup
        try:
            if controller and controller.excel is not None:
                controller.quit_excel()
        except Exception:
            pass

        try:
            if pythoncom:
                pythoncom.CoUninitialize()
        except Exception:
            pass

    def _dispatch(self, controller: ExcelController, task: _Task) -> Any:
        action = (task.action or "").strip().lower()

        if action == "launch":
            if controller.excel is None:
                controller.launch_new_instance()
                controller.set_excel_mode(controller.mode)
            return True

        if action == "quit":
            if controller.excel is not None:
                controller.quit_excel()
            return True

        if action == "set_mode":
            mode = (task.args[0] if task.args else task.kwargs.get("mode", "minimized"))
            controller.mode = str(mode).strip().lower() or "minimized"
            if controller.excel is not None:
                controller.set_excel_mode(controller.mode)
            return controller.mode

        if action == "show_10s":
            if controller.excel is None:
                controller.launch_new_instance()
            prev = controller.mode
            controller.show_excel_for_seconds(10)
            time.sleep(10)
            try:
                controller.set_excel_mode(prev)
            except Exception:
                pass
            return True

        if action == "run_pilot":
            # ----------------------------
            # Backward-compatible parsing:
            # - preferred: kwargs pilot_path/macro/args/excel_mode
            # - fallback: positional args (pilot_path, macro, args, excel_mode)
            # ----------------------------
            if "pilot_path" in task.kwargs:
                pilot_path = task.kwargs["pilot_path"]
                macro = task.kwargs["macro"]
                args = task.kwargs.get("args", [])
                excel_mode = task.kwargs.get("excel_mode", controller.mode)
            else:
                if len(task.args) < 2:
                    raise RuntimeError(
                        "run_pilot requires (pilot_path, macro, [args], [excel_mode]) or kwargs."
                    )
                pilot_path = task.args[0]
                macro = task.args[1]
                args = task.args[2] if len(task.args) >= 3 else []
                excel_mode = task.args[3] if len(task.args) >= 4 else controller.mode

            # Normalize args
            if args is None:
                args = []
            if isinstance(args, tuple):
                args = list(args)

            if controller.excel is None:
                controller.launch_new_instance()

            desired = str(excel_mode).strip().lower() or controller.mode

            # Hide main Excel window, keep dialogs/userforms visible (UIWatcher handles them)
            if desired in ("minimized", "hidden"):
                controller.set_excel_mode("hidden")
            else:
                controller.set_excel_mode("visible")

            wb_name = controller.open_or_activate_by_path(str(pilot_path))

            try:
                controller.run_macro(wb_name, str(macro), *list(args))
            finally:
                # Restore user preference after macro ends
                try:
                    controller.set_excel_mode(desired)
                except Exception:
                    pass

            return True

        raise RuntimeError(f"Unknown ExcelWorker action: {task.action}")
