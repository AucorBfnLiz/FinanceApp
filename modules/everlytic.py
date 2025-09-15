# modules/pettycash_page.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from modules.base_page import BasePage

class Everlytic(BasePage):
    def __init__(self, parent):
        steps = [
            "1. Select the cleaned petty cash Excel (header on row 4, data from row 7).",
            "2. Only rows with valid dates will be used.",
            "3. UseTax = Y only if VAT = Y.",
            "4. Module: gl→ general ledger, ar→ accounts receivable, ap→ accounts payable.",
            "5. Blank Reference becomes 'DEP'.",
        ]
        super().__init__(parent, "🧾 Petty Cash Import", steps)

        self.df = None
        self.build_ui()

    def build_ui(self):
        # Buttons
        control_frame = ttk.Frame(self.frame)
        control_frame.pack(pady=10)

        ttk.Button(control_frame, text="📁 Select Petty Cash File", command=self.load_file).grid(row=0, column=0, padx=10)
        ttk.Button(control_frame, text="⬇️ Export Processed", command=self.export_file).grid(row=0, column=1, padx=10)

        # Treeview
        self.tree = ttk.Treeview(self.frame, show="headings")
        self.tree.pack(fill="both", expand=True, pady=10)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            df = pd.read_excel(path, header=3).iloc[2:].reset_index(drop=True)
            df.columns = df.columns.str.strip()
            self.df = df.dropna(how="all")
            self.display_data(self.df)
            messagebox.showinfo("Loaded", "Petty Cash file loaded successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")

    def display_data(self, df):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="w")
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

    def export_file(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("No Data", "Load a file first.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx")
        if path:
            self.df.to_excel(path, index=False)
            messagebox.showinfo("Exported", f"Saved to {path}")
