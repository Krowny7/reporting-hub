from __future__ import annotations

import customtkinter as ctk

from ..gui.widgets import Card, btn_primary


def build_settings_page(app, parent) -> ctk.CTkFrame:
    page = ctk.CTkFrame(parent, fg_color="transparent")
    page.grid_columnconfigure(0, weight=1)

    card = Card(page, "Settings", "Saved in settings.json")
    card.grid(row=0, column=0, sticky="nsew")
    card.grid_columnconfigure(0, weight=1)

    btn_primary(card, "Save", command=app.on_save_settings, height=46).grid(
        row=3, column=0, padx=18, pady=(0, 18), sticky="ew"
    )
    return page
