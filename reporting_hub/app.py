from __future__ import annotations

import os
import ctypes
import time
from datetime import datetime

import customtkinter as ctk
from tkinter import filedialog

from .config.constants import (
    APP_TITLE,
    DEFAULT_PILOT_MACRO,
    SETTINGS_PATH,
    REPORT_TYPE_OPTIONS,
    DEFAULT_REPORT_TYPE,
    REPORT_TYPE_DEFAULT_MACROS,
)
from .config.io import load_settings, save_settings
from .config.models import MacroDefinition

from .excel.worker import ExcelWorker
from .gui.style import BG_APP, MUTED, TEXT, apply_app_style
from .gui.widgets import Card, ToastHost, btn_primary, btn_ghost
from .pages.update import build_update_page
from .pages.emails import build_emails_page
from .pages.settings import build_settings_page


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Load settings first (to apply appearance)
        self.settings = load_settings(SETTINGS_PATH)
        apply_app_style(self.settings.appearance or "Dark")

        # Active report type (weekly/monthly/quarterly/semiannual)
        self._active_report_type = (
            (self.settings.report_type or DEFAULT_REPORT_TYPE).strip().lower()
            or DEFAULT_REPORT_TYPE
        )

        self.configure(fg_color=BG_APP)
        self.title(APP_TITLE)
        self.minsize(1200, 760)
        self.geometry("1340x860")

        # maximize reliably AFTER window exists
        self.after(50, self._maximize_window_reliably)

        # Excel COM must run on a single dedicated thread.
        self.excel_worker: ExcelWorker | None = None
        self._running = False
        self._run_start_ts: float | None = None

        # Optional references (kept for future extensions)
        self._update_grid = None
        self._card_run = None
        self._card_side = None
        self._card_actions = None
        self._card_mini = None

        self._build_root()

        # Start the Excel worker AFTER UI widgets exist (toast/log uses Tk).
        self.excel_worker = ExcelWorker(self, ui_log=self.log, ui_toast=self.toast.show)

        self._apply_settings_to_widgets()

        # Start background worker (Excel COM thread)
        try:
            self.excel_worker.start()
            self.excel_worker.submit("set_mode", self.excel_mode.get())
        except Exception:
            pass

        # Graceful shutdown
        try:
            self.protocol("WM_DELETE_WINDOW", self.on_close)
        except Exception:
            pass

        self.toast.show("Ready.")

    def on_close(self):
        try:
            if self.excel_worker is not None:
                self.excel_worker.stop()
        except Exception:
            pass
        try:
            self.destroy()
        except Exception:
            pass

    def _maximize_window_reliably(self):
        try:
            self.update_idletasks()
            self.state("zoomed")
            self.update_idletasks()
            if str(self.state()).lower() == "zoomed":
                return
        except Exception:
            pass

        try:
            hwnd = self.winfo_id()
            ctypes.windll.user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
            return
        except Exception:
            pass

        try:
            w = self.winfo_screenwidth()
            h = self.winfo_screenheight()
            self.geometry(f"{w}x{h}+0+0")
        except Exception:
            pass

    # ---------- Layout ----------
    def _build_root(self):
        self.root = ctk.CTkFrame(self, corner_radius=0, fg_color=BG_APP)
        self.root.pack(fill="both", expand=True)

        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

        # Sidebar
        self.sidebar = ctk.CTkFrame(
            self.root,
            width=280,
            corner_radius=0,
            fg_color=("#FFFFFF", "#111112"),
            border_width=1,
            border_color=("#E5E5EA", "#232325"),
        )
        self.sidebar.grid(row=0, column=0, sticky="nsw")
        self.sidebar.grid_rowconfigure(50, weight=1)

        # Content
        self.content = ctk.CTkFrame(self.root, corner_radius=0, fg_color=BG_APP)
        self.content.grid(row=0, column=1, sticky="nsew")
        self.content.grid_rowconfigure(1, weight=1)
        self.content.grid_columnconfigure(0, weight=1)

        # Topbar
        self.topbar = ctk.CTkFrame(self.content, corner_radius=0, fg_color=BG_APP)
        self.topbar.grid(row=0, column=0, sticky="ew", padx=22, pady=(18, 10))
        self.topbar.grid_columnconfigure(1, weight=1)

        self.page_title = ctk.CTkLabel(
            self.topbar,
            text="Update",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=TEXT,
        )
        self.page_title.grid(row=0, column=0, sticky="w")

        self.quick_status = ctk.CTkLabel(self.topbar, text="Ready", text_color=MUTED)
        self.quick_status.grid(row=0, column=1, sticky="w", padx=(14, 0))

        self.appearance = ctk.CTkOptionMenu(
            self.topbar,
            values=["Dark", "Light", "System"],
            command=self.on_change_appearance,
            width=130,
        )
        self.appearance.set(self.settings.appearance or "Dark")
        self.appearance.grid(row=0, column=2, sticky="e", padx=(10, 0))

        # Pages
        self.pages = ctk.CTkFrame(self.content, corner_radius=0, fg_color=BG_APP)
        self.pages.grid(row=1, column=0, sticky="nsew", padx=22, pady=(0, 12))
        self.pages.grid_rowconfigure(0, weight=1)
        self.pages.grid_columnconfigure(0, weight=1)

        # Toasts
        self.toast = ToastHost(self.content)

        # Sidebar + pages
        self._build_sidebar()

        self.page_update = build_update_page(self, self.pages)
        self.page_emails = build_emails_page(self, self.pages)
        self.page_settings = build_settings_page(self, self.pages)

        self.show_page("update")

    def _build_sidebar(self):
        ctk.CTkLabel(
            self.sidebar,
            text="REPORTING HUB",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=MUTED,
        ).grid(row=0, column=0, padx=18, pady=(18, 6), sticky="w")

        ctk.CTkLabel(
            self.sidebar,
            text="Automation",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=TEXT,
        ).grid(row=1, column=0, padx=18, pady=(0, 18), sticky="w")

        btn_primary(self.sidebar, "Update", command=lambda: self.show_page("update"), height=42).grid(
            row=2, column=0, padx=18, pady=(0, 10), sticky="ew"
        )
        btn_ghost(self.sidebar, "Emails (UI)", command=lambda: self.show_page("emails"), height=42).grid(
            row=3, column=0, padx=18, pady=0, sticky="ew"
        )
        btn_ghost(self.sidebar, "Settings", command=lambda: self.show_page("settings"), height=42).grid(
            row=4, column=0, padx=18, pady=10, sticky="ew"
        )

        excel_card = Card(self.sidebar, "Excel", "Discret + dialogues visibles")
        excel_card.grid(row=10, column=0, padx=18, pady=(18, 12), sticky="ew")

        self.excel_mode = ctk.StringVar(value="minimized")
        ctk.CTkOptionMenu(
            excel_card,
            values=["minimized", "hidden", "visible"],
            variable=self.excel_mode,
            command=self.on_change_excel_mode,
            corner_radius=18,
        ).grid(row=2, column=0, padx=18, pady=(0, 12), sticky="ew")

        row = ctk.CTkFrame(excel_card, fg_color="transparent")
        row.grid(row=3, column=0, padx=18, pady=(0, 16), sticky="ew")
        row.grid_columnconfigure((0, 1), weight=1)

        btn_primary(row, "Launch", command=self.on_launch_excel, height=40).grid(
            row=0, column=0, padx=(0, 8), sticky="ew"
        )
        btn_ghost(row, "Quit", command=self.on_quit_excel, height=40).grid(
            row=0, column=1, padx=(8, 0), sticky="ew"
        )

        btn_ghost(excel_card, "Show 10s", command=self.on_show_excel_10s, height=40).grid(
            row=4, column=0, padx=18, pady=(0, 18), sticky="ew"
        )

        # Spacer to push logs to bottom
        spacer = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        spacer.grid(row=50, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(50, weight=1)

        logs_card = Card(self.sidebar, "Logs", "Execution output")
        logs_card.grid(row=60, column=0, padx=18, pady=(0, 18), sticky="ew")

        self.logbox = ctk.CTkTextbox(
            logs_card,
            height=140,
            corner_radius=18,
            fg_color=("#F7F7FA", "#121214"),
            border_width=1,
            border_color=("#E5E5EA", "#232325"),
            wrap="none",
        )
        self.logbox.grid(row=2, column=0, padx=18, pady=(10, 18), sticky="ew")

    # ---------- Pages ----------
    def _clear_pages(self):
        for child in self.pages.winfo_children():
            child.grid_forget()

    def show_page(self, key: str) -> None:
        self._clear_pages()
        if key == "update":
            self.page_title.configure(text="Update")
            self.page_update.grid(row=0, column=0, sticky="nsew")
        elif key == "emails":
            self.page_title.configure(text="Emails")
            self.page_emails.grid(row=0, column=0, sticky="nsew")
        else:
            self.page_title.configure(text="Settings")
            self.page_settings.grid(row=0, column=0, sticky="nsew")

    # ---------- Profiles ----------
    def _report_type_label(self, key: str) -> str:
        k = (key or DEFAULT_REPORT_TYPE).strip().lower() or DEFAULT_REPORT_TYPE
        return k[:1].upper() + k[1:]

    def _report_type_key(self, label: str) -> str:
        k = (label or "").strip().lower()
        allowed = {v.strip().lower() for v in REPORT_TYPE_OPTIONS}
        return k if k in allowed else DEFAULT_REPORT_TYPE

    def _default_macro_for(self, report_key: str) -> str:
        key = (report_key or DEFAULT_REPORT_TYPE).strip().lower() or DEFAULT_REPORT_TYPE
        return REPORT_TYPE_DEFAULT_MACROS.get(key, DEFAULT_PILOT_MACRO)

    def _get_profile(self, report_key: str) -> MacroDefinition:
        key = (report_key or DEFAULT_REPORT_TYPE).strip().lower() or DEFAULT_REPORT_TYPE
        prof = self.settings.macros.get(key)
        if prof is None:
            prof = MacroDefinition(
                label=self._report_type_label(key),
                workbook_path="",
                macro=self._default_macro_for(key),
                args="",
            )
            self.settings.macros[key] = prof
        # Auto-fix: if profile has empty macro, set a sensible default
        if not (prof.macro or "").strip():
            prof.macro = self._default_macro_for(key)
        return prof

    def _set_entry(self, entry, value: str) -> None:
        if entry is None:
            return
        try:
            entry.delete(0, "end")
            if value:
                entry.insert(0, value)
        except Exception:
            pass

    def _persist_profile(self, report_key: str) -> None:
        key = (report_key or DEFAULT_REPORT_TYPE).strip().lower() or DEFAULT_REPORT_TYPE
        prof = self._get_profile(key)

        prof.workbook_path = self.pilot_path_entry.get().strip() if getattr(self, "pilot_path_entry", None) else ""
        macro_in = self.pilot_macro_entry.get().strip() if getattr(self, "pilot_macro_entry", None) else ""
        prof.macro = macro_in or self._default_macro_for(key)
        prof.args = self.pilot_args_entry.get().strip() if getattr(self, "pilot_args_entry", None) else ""
        prof.label = self._report_type_label(key)

        self.settings.macros[key] = prof

    def _apply_settings_to_widgets(self) -> None:
        mode = (self.settings.excel_mode or "minimized").strip().lower()
        self.excel_mode.set(mode)

        self._active_report_type = (
            (self.settings.report_type or DEFAULT_REPORT_TYPE).strip().lower()
            or DEFAULT_REPORT_TYPE
        )

        if getattr(self, "report_type_var", None) is not None:
            self.report_type_var.set(self._report_type_label(self._active_report_type))

        prof = self._get_profile(self._active_report_type)
        self._set_entry(getattr(self, "pilot_path_entry", None), prof.workbook_path)
        self._set_entry(getattr(self, "pilot_macro_entry", None), prof.macro or self._default_macro_for(self._active_report_type))
        self._set_entry(getattr(self, "pilot_args_entry", None), prof.args or "")

    def _persist_settings_from_widgets(self) -> None:
        self.settings.appearance = self.appearance.get().strip() or "Dark"
        self.settings.excel_mode = self.excel_mode.get().strip().lower() or "minimized"

        label = (
            self.report_type_var.get()
            if getattr(self, "report_type_var", None) is not None
            else self._report_type_label(self._active_report_type)
        )
        self._active_report_type = self._report_type_key(label)
        self.settings.report_type = self._active_report_type

        self._persist_profile(self._active_report_type)
        save_settings(SETTINGS_PATH, self.settings)

    def on_save_settings(self):
        """Used by Settings page button."""
        self._persist_settings_from_widgets()
        self.toast.show("Saved.")

    def on_change_appearance(self, mode: str):
        try:
            ctk.set_appearance_mode(mode)
        except Exception:
            pass
        self.settings.appearance = mode
        save_settings(SETTINGS_PATH, self.settings)

    def on_change_report_type(self, value: str):
        # Save current inputs before switching
        try:
            self._persist_profile(self._active_report_type)
        except Exception:
            pass

        new_key = self._report_type_key(value)
        self._active_report_type = new_key
        self.settings.report_type = new_key

        prof = self._get_profile(new_key)

        # If macro was never changed and is still the monthly default, auto-switch to frequency default
        if (prof.macro or "").strip() in ("", DEFAULT_PILOT_MACRO) and new_key != "monthly":
            prof.macro = self._default_macro_for(new_key)
            self.settings.macros[new_key] = prof

        self._set_entry(getattr(self, "pilot_path_entry", None), prof.workbook_path)
        self._set_entry(getattr(self, "pilot_macro_entry", None), prof.macro or self._default_macro_for(new_key))
        self._set_entry(getattr(self, "pilot_args_entry", None), prof.args or "")

        if getattr(self, "report_type_var", None) is not None:
            self.report_type_var.set(self._report_type_label(new_key))

        save_settings(SETTINGS_PATH, self.settings)
        self.toast.show(f"{self._report_type_label(new_key)} selected.")

    # ---------- Logging ----------
    def log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        line = f"[{ts}] {msg}\n"
        try:
            self.logbox.insert("end", line)
            self.logbox.see("end")
        except Exception:
            pass
        try:
            self.quick_status.configure(text=msg, text_color=MUTED)
        except Exception:
            pass

    # ---------- Long-running jobs UX ----------
    def _set_running(self, running: bool):
        self._running = bool(running)
        if not hasattr(self, "progress") or not hasattr(self, "run_btn"):
            return

        if self._running:
            self._run_start_ts = time.time()
            try:
                self.progress.configure(mode="indeterminate")
                self.progress.start()
            except Exception:
                try:
                    self.progress.set(0.12)
                except Exception:
                    pass
            try:
                self.run_btn.configure(state="disabled")
            except Exception:
                pass
            self.after(250, self._tick_running)
        else:
            try:
                self.progress.stop()
                self.progress.configure(mode="determinate")
                self.progress.set(0.0)
            except Exception:
                try:
                    self.progress.set(0.0)
                except Exception:
                    pass
            try:
                self.run_btn.configure(state="normal")
            except Exception:
                pass

    def _tick_running(self):
        if not self._running:
            return
        try:
            elapsed = int(time.time() - float(self._run_start_ts or time.time()))
            mm, ss = divmod(elapsed, 60)
            hh, mm = divmod(mm, 60)
            label = f"Running… {hh:02d}:{mm:02d}:{ss:02d}" if hh else f"Running… {mm:02d}:{ss:02d}"
            self.quick_status.configure(text=label, text_color=MUTED)
        except Exception:
            pass
        self.after(1000, self._tick_running)

    # ---------- Excel controls ----------
    def on_change_excel_mode(self, value: str):
        self._persist_settings_from_widgets()
        if self.excel_worker is None:
            self.toast.show(f"Mode saved: {value}")
            return
        self.excel_worker.submit(
            "set_mode",
            value,
            on_ok=lambda _r: self.toast.show(f"Excel mode: {value}"),
            on_err=lambda e: self.toast.show(str(e)),
        )

    def on_launch_excel(self):
        if self.excel_worker is None:
            self.toast.show("Excel worker not ready.")
            return
        self.excel_worker.submit(
            "launch",
            self.excel_mode.get() if hasattr(self, "excel_mode") else "minimized",
            on_ok=lambda _r: self.toast.show("Excel ready."),
            on_err=lambda e: self.toast.show(f"Excel error: {e}"),
        )

    def on_quit_excel(self):
        if self.excel_worker is None:
            return
        self.excel_worker.submit(
            "quit",
            on_ok=lambda _r: self.toast.show("Excel closed."),
            on_err=lambda e: self.toast.show(f"Close error: {e}"),
        )

    def on_show_excel_10s(self):
        if self.excel_worker is None:
            return
        self.excel_worker.submit(
            "show_10s",
            10,
            on_ok=lambda _r: self.toast.show("Showing Excel (10s)."),
            on_err=lambda e: self.toast.show(str(e)),
        )

    # ---------- Update flow ----------
    def on_pick_pilot(self):
        path = filedialog.askopenfilename(
            title="Choose pilot workbook",
            filetypes=[("Excel macro-enabled", "*.xlsm"), ("Excel", "*.xlsx;*.xlsb;*.xls")],
        )
        if not path:
            return
        self.pilot_path_entry.delete(0, "end")
        self.pilot_path_entry.insert(0, path)
        self._persist_settings_from_widgets()
        self.toast.show("Pilot selected.")

    def on_run_pilot(self):
        # Persist current UI state
        self._persist_settings_from_widgets()

        pilot_path = self.pilot_path_entry.get().strip()
        macro = (self.pilot_macro_entry.get().strip() or self._default_macro_for(self._active_report_type))
        raw_args = self.pilot_args_entry.get().strip()

        if not pilot_path or not os.path.exists(pilot_path):
            self.toast.show("Pilot not found.")
            return

        args = [a.strip() for a in raw_args.split(";") if a.strip()] if raw_args else []

        if self.excel_worker is None:
            self.toast.show("Excel worker not ready.")
            return

        self._set_running(True)
        self.toast.show("Running…")

        excel_mode = self.excel_mode.get() if hasattr(self, "excel_mode") else "minimized"

        def ok(_result: object) -> None:
            self._set_running(False)
            self.toast.show("Done.")
            self.log("Done.")

        def err(e: BaseException) -> None:
            self._set_running(False)
            self.toast.show("Error (see log).")
            self.log(str(e))

        self.excel_worker.submit(
            "run_pilot",
            pilot_path,
            macro,
            args,
            excel_mode,
            on_ok=ok,
            on_err=err,
        )
