from __future__ import annotations

import customtkinter as ctk

from .style import BORDER, CARD, BTN_GHOST_HOVER, BTN_MAIN_BG, BTN_MAIN_HOVER, BTN_MAIN_TEXT, TEXT, MUTED, font


class Card(ctk.CTkFrame):
    def __init__(self, master, title: str, subtitle: str = "", **kwargs):
        super().__init__(
            master,
            corner_radius=22,
            border_width=1,
            fg_color=CARD,
            border_color=BORDER,
            **kwargs,
        )
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text=title, font=font(16, "bold"), text_color=TEXT).grid(
            row=0, column=0, padx=18, pady=(18, 2), sticky="w"
        )
        if subtitle:
            ctk.CTkLabel(self, text=subtitle, text_color=MUTED).grid(
                row=1, column=0, padx=18, pady=(0, 12), sticky="w"
            )


class ToastHost(ctk.CTkFrame):
    def __init__(self, master):
        super().__init__(
            master,
            fg_color="transparent",
            corner_radius=0,
            border_width=0,
            width=1,
            height=1,
        )
        self.pack_propagate(False)
        self.grid_propagate(False)
        self.place(relx=1.0, rely=0.0, anchor="ne", x=-18, y=18)

    def show(self, text: str, ttl_ms: int = 3200) -> None:
        toast = ctk.CTkFrame(self, corner_radius=18, border_width=1, fg_color=CARD, border_color=BORDER)
        toast.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            toast,
            text=text,
            wraplength=520,
            font=font(13, "normal"),
            text_color=TEXT,
        ).grid(row=0, column=0, padx=16, pady=12, sticky="w")

        toast.pack(fill="x", pady=8)

        def kill():
            try:
                toast.destroy()
            except Exception:
                pass

        toast.after(ttl_ms, kill)


def btn_primary(master, text: str, command=None, height: int = 44) -> ctk.CTkButton:
    return ctk.CTkButton(
        master,
        text=text,
        command=command,
        height=height,
        corner_radius=18,
        fg_color=BTN_MAIN_BG,
        hover_color=BTN_MAIN_HOVER,
        text_color=BTN_MAIN_TEXT,
    )


def btn_ghost(master, text: str, command=None, height: int = 40) -> ctk.CTkButton:
    return ctk.CTkButton(
        master,
        text=text,
        command=command,
        height=height,
        corner_radius=18,
        fg_color="transparent",
        hover_color=BTN_GHOST_HOVER,
        border_width=1,
        border_color=BORDER,
        text_color=TEXT,
    )
