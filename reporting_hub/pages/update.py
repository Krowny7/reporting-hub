from __future__ import annotations

import customtkinter as ctk

from ..config.constants import DEFAULT_PILOT_MACRO, REPORT_TYPE_OPTIONS
from ..gui.style import BG_APP, BORDER, FIELD, TEXT
from ..gui.widgets import Card, btn_primary, btn_ghost


def build_update_page(app, parent) -> ctk.CTkFrame:
    """Update page (pilot workbook + macro runner)."""
    page = ctk.CTkFrame(parent, fg_color=BG_APP)
    page.grid_rowconfigure(0, weight=1)
    page.grid_columnconfigure(0, weight=1)

    grid = ctk.CTkFrame(page, fg_color=BG_APP)
    grid.grid(row=0, column=0, sticky="nsew")

    app._update_grid = grid

    grid.grid_columnconfigure(0, weight=1, uniform="col")
    grid.grid_columnconfigure(1, weight=1, uniform="col")

    # Row 0 = main card (full width)
    # Row 1 = secondary cards
    # Row 2 = spacer (absorbs remaining height)
    grid.grid_rowconfigure(0, weight=0)
    grid.grid_rowconfigure(1, weight=0)
    grid.grid_rowconfigure(2, weight=1)

    ROW_GAP = 26

    # --- Top: Run (full width) ---
    run_card = Card(grid, "Monthly update", "Select frequency → open pilot + run macro")
    run_card.grid(row=0, column=0, columnspan=2, padx=0, pady=(0, ROW_GAP), sticky="new")
    run_card.grid_columnconfigure(0, weight=1)
    app._card_run = run_card

    # Report type selector
    app.report_type_var = ctk.StringVar(value="Monthly")
    app.report_type_menu = ctk.CTkOptionMenu(
        run_card,
        values=REPORT_TYPE_OPTIONS,
        variable=app.report_type_var,
        command=app.on_change_report_type,
        corner_radius=18,
        height=40,
    )
    app.report_type_menu.grid(row=2, column=0, padx=18, pady=(0, 12), sticky="ew")

    app.pilot_path_entry = ctk.CTkEntry(
        run_card,
        placeholder_text="Pilot workbook (.xlsm) — OneDrive synced path",
        fg_color=FIELD,
        border_color=BORDER,
        text_color=TEXT,
        corner_radius=18,
        height=40,
    )
    app.pilot_path_entry.grid(row=3, column=0, padx=18, pady=(0, 12), sticky="ew")

    btn_ghost(run_card, "Choose pilot", command=app.on_pick_pilot, height=40).grid(
        row=4, column=0, padx=18, pady=(0, 12), sticky="ew"
    )

    app.pilot_macro_entry = ctk.CTkEntry(
        run_card,
        placeholder_text=f"Macro (default: {DEFAULT_PILOT_MACRO})",
        fg_color=FIELD,
        border_color=BORDER,
        text_color=TEXT,
        corner_radius=18,
        height=40,
    )
    app.pilot_macro_entry.grid(row=5, column=0, padx=18, pady=(0, 12), sticky="ew")

    app.pilot_args_entry = ctk.CTkEntry(
        run_card,
        placeholder_text="Optional args (separated by ;)",
        fg_color=FIELD,
        border_color=BORDER,
        text_color=TEXT,
        corner_radius=18,
        height=40,
    )
    app.pilot_args_entry.grid(row=6, column=0, padx=18, pady=(0, 14), sticky="ew")

    app.progress = ctk.CTkProgressBar(run_card, corner_radius=999)
    app.progress.set(0)
    app.progress.grid(row=7, column=0, padx=18, pady=(0, 12), sticky="ew")

    app.run_btn = btn_primary(run_card, "Run", command=app.on_run_pilot, height=46)
    app.run_btn.grid(row=8, column=0, padx=18, pady=(0, 18), sticky="ew")

    # --- Removed: Reliability card ---
    app._card_side = None

    # --- Bottom row: Actions (left) ---
    actions = Card(grid, "Reporting tools", "UI placeholders (we’ll connect later)")
    actions.grid(row=1, column=0, padx=(0, 12), pady=(0, 0), sticky="new")
    actions.grid_columnconfigure((0, 1), weight=1)
    app._card_actions = actions

    btn_ghost(actions, "Export PDFs (soon)", height=42).grid(row=2, column=0, padx=18, pady=18, sticky="ew")
    btn_ghost(actions, "Send emails (soon)", command=lambda: app.show_page("emails"), height=42).grid(
        row=2, column=1, padx=18, pady=18, sticky="ew"
    )

    # --- Bottom row: Mini card (right) ---
    mini = Card(grid, "Recent events", "See bottom log for details")
    mini.grid(row=1, column=1, padx=(12, 0), pady=(0, 0), sticky="new")
    app._card_mini = mini

    # Spacer row (absorbs extra height so cards don't stretch)
    ctk.CTkFrame(grid, fg_color=BG_APP).grid(row=2, column=0, columnspan=2, sticky="nsew")

    return page
