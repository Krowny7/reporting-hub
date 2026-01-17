from __future__ import annotations

import customtkinter as ctk


def C(light: str, dark: str):
    """Return a (light, dark) tuple used by CustomTkinter."""
    return (light, dark)


# Surfaces
BG_APP = C("#F5F5F7", "#0B0B0C")
SIDEBAR = C("#FFFFFF", "#111112")
CARD = C("#FFFFFF", "#141415")
FIELD = C("#FFFFFF", "#101011")
BORDER = C("#E5E5EA", "#232325")

# Text
TEXT = C("#1D1D1F", "#F5F5F7")
MUTED = C("#6E6E73", "#A1A1A6")

# Buttons
BTN_MAIN_BG = C("#1D1D1F", "#F5F5F7")
BTN_MAIN_TEXT = C("#FFFFFF", "#0B0B0C")
BTN_MAIN_HOVER = C("#2C2C2E", "#E9E9EC")

BTN_GHOST_BG = "transparent"
BTN_GHOST_HOVER = C("#F2F2F3", "#1C1C1E")


def apply_app_style(appearance: str = "Dark") -> None:
    """Apply global CTk style."""
    ctk.set_appearance_mode(appearance)
    # Keep CTk internals stable
    ctk.set_default_color_theme("blue")


def font(size: int = 13, weight: str = "normal"):
    try:
        return ctk.CTkFont(size=size, weight=weight)
    except Exception:
        return None
