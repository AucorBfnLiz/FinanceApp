# modules/base_page.py
import tkinter as tk
from tkinter import ttk

class BasePage:
    def __init__(self, parent, title, steps):
        self.parent = parent
        self.frame = ttk.Frame(parent, padding=20)
        self.frame.pack(fill="both", expand=True)

        ttk.Label(self.frame, text=title, font=("Segoe UI", 14, "bold")).pack(pady=(0, 10))

        # How it works
        info_frame = ttk.LabelFrame(self.frame, text="How it works", padding=10)
        info_frame.pack(fill="x", pady=10)
        for step in steps:
            ttk.Label(info_frame, text=step, font=("Segoe UI", 10)).pack(anchor="w", padx=10, pady=1)
