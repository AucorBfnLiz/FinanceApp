import tkinter as tk
from tkinter import ttk

# === Module imports ===
from modules.payment_requisition import PaymentRequisitionApp
from modules.import9500 import DepositImportApp
from modules.expenses import GLExtractorApp
from modules.bidmasterimport import BidmasterSalesApp
from modules.compare9500 import Compare9500App
from modules.petty_cash import PettyCashApp
from modules.everlytic import Everlytic
# from modules.creditors import CreditorsApp  # Uncomment once ready


def build_app():
    root = tk.Tk()
    root.title("Finance Automation Toolkit")
    root.geometry("900x650")

    # ---- Window/grid defaults for proper resizing ----
    root.rowconfigure(0, weight=1)
    root.columnconfigure(0, weight=1)

    # ---- Styles / Theme ----
    style = ttk.Style()
    # You can try: "clam", "alt", "default", "vista", "xpnative" (Windows)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure("Title.TLabel", font=("Arial", 18, "bold"))
    style.configure("TNotebook.Tab", font=("Arial", 11, "bold"), padding=(16, 8))
    style.configure("TNotebook", tabposition="n")

    # ---- Notebook ----
    notebook = ttk.Notebook(root)
    notebook.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

    # ---- Welcome tab (static) ----
    welcome_frame = ttk.Frame(notebook)
    ttk.Label(welcome_frame, text="Welcome to the Finance Toolkit", style="Title.TLabel").pack(pady=24)
    notebook.add(welcome_frame, text="🏁 Welcome")

    # ---- Tab registry: title → class ----
    tabs = [
        ("💼 EWallet & Petty Cash", PettyCashApp),
        ("📊 Compare 9500",       Compare9500App),
        ("🏦 Import Deposits",     DepositImportApp),
        ("💸 Requisitions",        PaymentRequisitionApp),
        ("🧾 Bidmaster",          BidmasterSalesApp),
        ("📂 GL Extractor",       GLExtractorApp),
        ("📑 Everlytic",          Everlytic),
        # ("🧾 Creditors",        CreditorsApp),
    ]

    # ---- Create each tab from the registry ----
    for title, AppClass in tabs:
        frame = ttk.Frame(notebook)
        # Let your module build its own layout inside the frame
        AppClass(frame)
        notebook.add(frame, text=title)

    return root


if __name__ == "__main__":
    app = build_app()
    app.mainloop()
