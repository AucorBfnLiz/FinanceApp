# modules/theme.py
import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont

# ---------- Palette ----------
def get_palette(root) -> dict:
    """
    Return shared colors. Keep keys stable so screens can rely on them.
    """
    return {
        "BG":       "#ffffff",
        "FG":       "#111827",
        "MUTED":    "#6b7280",
        "PRIMARY":  "#2563eb",
        "PRIMARY_FG": "#ffffff",
        "BORDER":   "#e5e7eb",
        "STRIPE":   "#f9fafb",   # used for Treeview alternating rows
        "CARD_BG":  "#ffffff",
        "CARD_BORDER": "#e5e7eb",
    }

# ---------- Fonts & Styles ----------
def _ensure_named_font(root, name, **cfg):
    try:
        f = tkfont.nametofont(name)
        # If it exists, keep it (don’t override user-changed sizes at runtime)
        return f
    except tk.TclError:
        return tkfont.Font(root, name=name, **cfg)

def init_styles(root: tk.Misc) -> None:
    """
    Idempotent. Call once early (e.g., in app.py after creating the root).
    Sets up named fonts and ttk styles used across modules.
    """
    pal = get_palette(root)

    # Named fonts (referenced by name so we can tweak later at runtime)
    _ensure_named_font(root, "App.Base",   family="Segoe UI", size=9)
    _ensure_named_font(root, "App.Small",  family="Segoe UI", size=8)
    _ensure_named_font(root, "App.Title",  family="Segoe UI", size=14, weight="bold")
    _ensure_named_font(root, "App.Subtle", family="Segoe UI", size=9)

    style = ttk.Style(root)
    # Use a predictable theme
    try:
        if style.theme_use() not in ("clam", "vista", "xpnative", "alt", "default"):
            style.theme_use("clam")
    except tk.TclError:
        pass

    # Global fallback for non-ttk widgets
    root.option_add("*Font", "App.Base")
    root.option_add("*foreground", pal["FG"])
    root.option_add("*Background", pal["BG"])

    # Base ttk
    style.configure("TLabel", font="App.Base", foreground=pal["FG"], background=pal["BG"])
    style.configure("TButton", font="App.Base")
    style.configure("TEntry", font="App.Base")
    style.configure("TCombobox", font="App.Base")
    style.configure("TLabelframe.Label", font="App.Base", foreground=pal["FG"])
    style.configure("TNotebook.Tab", font="App.Base")

    # Custom variants used in your modules
    style.configure("Title.TLabel", font="App.Title", foreground=pal["FG"], background=pal["BG"])
    style.configure("Subtle.TLabel", font="App.Subtle", foreground=pal["MUTED"], background=pal["BG"])

    # Primary button (keep it simple/cross‑platform)
    style.configure(
        "Primary.TButton",
        font="App.Base",
        padding=(10, 6),
        foreground=pal["PRIMARY_FG"],
        background=pal["PRIMARY"],
        borderwidth=1,
    )
    style.map(
        "Primary.TButton",
        background=[("active", pal["PRIMARY"]), ("pressed", pal["PRIMARY"])],
        foreground=[("disabled", "#cbd5e1")]
    )

    # Card frame
    style.configure(
        "Card.TFrame",
        background=pal["CARD_BG"],
        bordercolor=pal["CARD_BORDER"],
        relief="groove"
    )

    # Treeview
    style.configure(
        "Modern.Treeview",
        font="App.Base",
        rowheight=24,
        background=pal["BG"],
        fieldbackground=pal["BG"],
        foreground=pal["FG"],
        bordercolor=pal["BORDER"],
        lightcolor=pal["BORDER"],
        darkcolor=pal["BORDER"],
    )
    style.configure("Modern.Treeview.Heading", font="App.Base")

# ---------- Small UI helper ----------
def make_card(parent, title: str, theme: dict) -> ttk.Frame:
    """
    Create a simple 'card' with a header; returns the inner body frame
    so the caller can pack/grid their own widgets inside it.
    """
    outer = ttk.Frame(parent, style="Card.TFrame", padding=12)
    # Prefer grid if parent uses grid; else fall back to pack
    try:
        outer.grid_configure
        outer.grid(column=0, row=ttk.Frame(parent).grid_info().get("row", 0), sticky="ew", padx=0, pady=(0, 8))
    except Exception:
        outer.pack(fill="x", pady=(0, 8))

    ttk.Label(outer, text=title, style="Title.TLabel").pack(anchor="w", pady=(0, 6))
    body = ttk.Frame(outer)  # inherits background via style
    body.pack(fill="x")
    return body
