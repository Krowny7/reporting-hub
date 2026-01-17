from __future__ import annotations

import customtkinter as ctk

from ..config.constants import DEFAULT_PILOT_MACRO
from ..gui.style import BG_APP, BORDER, FIELD, TEXT, MUTED
from ..gui.widgets import Card, btn_primary, btn_ghost


def build_update_page(app, parent) -> ctk.CTkFrame:
    """Update page (pilot workbook + macro runner)."""
    # Avoid "transparent" containers: on some Windows + DPI combinations,
    # CTk canvas redraw order can create visual overlap artifacts.
    page = ctk.CTkFrame(parent, fg_color=BG_APP)
    page.grid_rowconfigure(0, weight=1)
    page.grid_columnconfigure(0, weight=1)

    grid = ctk.CTkFrame(page, fg_color=BG_APP)
    grid.grid(row=0, column=0, sticky="nsew")

    # Expose for row minsize stabilizer
    app._update_grid = grid

    grid.grid_columnconfigure(0, weight=1, uniform="col")
    grid.grid_columnconfigure(1, weight=1, uniform="col")

    # IMPORTANT (no scrollbar):
    # Avoid hard "minsize" constraints. On high DPI or smaller screens,
    # a forced minimum height makes the page content larger than its viewport.
    # Tk/CTk does not clip overflow, so widgets can visually draw over the log panel.
    # Let the grid be responsive instead (weights only).
    grid.grid_rowconfigure(0, weight=3)
    grid.grid_rowconfigure(1, weight=2)

    # Card radius = 22 in widgets.py. If the vertical gap between rows is too small,
    # rounded corners + borders can look like they overlap.
    ROW_GAP = 26

    run_card = Card(grid, "Monthly update", "One click → open pilot + run macro")
    run_card.grid(row=0, column=0, padx=(0, 12), pady=(0, ROW_GAP), sticky="nsew")
    run_card.grid_columnconfigure(0, weight=1)
    app._card_run = run_card

    app.pilot_path_entry = ctk.CTkEntry(
        run_card,
        placeholder_text="Pilot workbook (.xlsm) — OneDrive synced path",
        fg_color=FIELD,
        border_color=BORDER,
        text_color=TEXT,
        corner_radius=18,
        height=40,
    )
    app.pilot_path_entry.grid(row=2, column=0, padx=18, pady=(0, 12), sticky="ew")

    btn_ghost(run_card, "Choose pilot", command=app.on_pick_pilot, height=40).grid(
        row=3, column=0, padx=18, pady=(0, 12), sticky="ew"
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
    app.pilot_macro_entry.grid(row=4, column=0, padx=18, pady=(0, 12), sticky="ew")

    app.pilot_args_entry = ctk.CTkEntry(
        run_card,
        placeholder_text="Optional args (separated by ;)",
        fg_color=FIELD,
        border_color=BORDER,
        text_color=TEXT,
        corner_radius=18,
        height=40,
    )
    app.pilot_args_entry.grid(row=5, column=0, padx=18, pady=(0, 14), sticky="ew")

    app.progress = ctk.CTkProgressBar(run_card, corner_radius=999)
    app.progress.set(0)
    app.progress.grid(row=6, column=0, padx=18, pady=(0, 12), sticky="ew")

    app.run_btn = btn_primary(run_card, "Run", command=app.on_run_pilot, height=46)
    app.run_btn.grid(row=7, column=0, padx=18, pady=(0, 18), sticky="ew")

    side = Card(grid, "Reliability", "SharePoint sync OK — no auth prompt expected")
    side.grid(row=0, column=1, padx=(12, 0), pady=(0, ROW_GAP), sticky="nsew")
    app._card_side = side

    ctk.CTkLabel(
        side,
        text="• Recommended: minimized\n• Excel dialogs visible\n• Next: run profiles",
        justify="left",
        text_color=MUTED,
    ).grid(row=2, column=0, padx=18, pady=(0, 18), sticky="nw")

    actions = Card(grid, "Reporting tools", "UI placeholders (we’ll connect later)")
    actions.grid(row=1, column=0, padx=(0, 12), pady=(0, 0), sticky="nsew")
    actions.grid_columnconfigure((0, 1), weight=1)
    app._card_actions = actions

    btn_ghost(actions, "Export PDFs (soon)", height=42).grid(row=2, column=0, padx=18, pady=18, sticky="ew")
    btn_ghost(actions, "Send emails (soon)", command=lambda: app.show_page("emails"), height=42).grid(
        row=2, column=1, padx=18, pady=18, sticky="ew"
    )

    mini = Card(grid, "Recent events", "See bottom log for details")
    mini.grid(row=1, column=1, padx=(12, 0), pady=(0, 0), sticky="nsew")
    app._card_mini = mini

    return page
