from __future__ import annotations

import customtkinter as ctk

from ..gui.style import BORDER, FIELD, TEXT
from ..gui.widgets import Card, btn_primary, btn_ghost


def build_emails_page(app, parent) -> ctk.CTkFrame:
    page = ctk.CTkFrame(parent, fg_color="transparent")
    page.grid_columnconfigure(0, weight=1)

    card = Card(page, "Emails", "UI only (Outlook integration later)")
    card.grid(row=0, column=0, sticky="nsew")
    card.grid_columnconfigure(0, weight=1)
    card.grid_rowconfigure(4, weight=1)

    app.email_subject = ctk.CTkEntry(
        card,
        placeholder_text="Subject…",
        fg_color=FIELD,
        border_color=BORDER,
        text_color=TEXT,
        corner_radius=18,
        height=40,
    )
    app.email_subject.grid(row=2, column=0, padx=18, pady=(0, 12), sticky="ew")

    app.email_to = ctk.CTkEntry(
        card,
        placeholder_text="Recipients (a@b.com; c@d.com)…",
        fg_color=FIELD,
        border_color=BORDER,
        text_color=TEXT,
        corner_radius=18,
        height=40,
    )
    app.email_to.grid(row=3, column=0, padx=18, pady=(0, 12), sticky="ew")

    app.email_body = ctk.CTkTextbox(
        card,
        height=260,
        fg_color=FIELD,
        border_color=BORDER,
        text_color=TEXT,
        corner_radius=18,
    )
    app.email_body.grid(row=4, column=0, padx=18, pady=(0, 18), sticky="nsew")

    row = ctk.CTkFrame(card, fg_color="transparent")
    row.grid(row=5, column=0, padx=18, pady=(0, 18), sticky="ew")
    row.grid_columnconfigure((0, 1, 2), weight=1)

    btn_ghost(row, "Preview (soon)", height=42).grid(row=0, column=0, padx=6, sticky="ew")
    btn_ghost(row, "Attach PDF (soon)", height=42).grid(row=0, column=1, padx=6, sticky="ew")
    btn_primary(row, "Send (soon)", height=42).grid(row=0, column=2, padx=6, sticky="ew")

    return page
