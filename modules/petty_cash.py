# modules/pettycash_page.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

from modules.base_page import BasePage
from modules.petty_transform import transform_petty_or_ewallet

class PettyCashApp(BasePage):
    def __init__(self, parent):
        steps = [
            "1. Select the cleaned petty cash Excel (header on row 4, data from row 7).",
            "2. Only rows with valid dates will be used.",
            "3. UseTax = Y only if VAT = Y.",
            "4. Module: gl→0, ar→1, ap→2.",
            "5. Blank Reference becomes 'DEP'.",
        ]
        super().__init__(parent, "🧾 Petty Cash Import", steps)

        self.src_df = None       # raw post-header/pre-clean df (after dropping 2 rows)
        self.proc_df = None      # transformed df (ready for export)

        self._build_ui()

    # ---------- UI ----------
    def _build_ui(self):
        # Controls
        controls = ttk.Frame(self.frame)
        controls.pack(pady=8, fill="x")

        ttk.Button(controls, text="📁 Select Petty Cash File", command=self.load_file).pack(side="left", padx=6)
        ttk.Button(controls, text="⬇️ Export to Excel", command=lambda: self.export_file(kind="xlsx")).pack(side="left", padx=6)
        ttk.Button(controls, text="⬇️ Export to CSV", command=lambda: self.export_file(kind="csv")).pack(side="left", padx=6)

        # Table
        table_wrap = ttk.Frame(self.frame)
        table_wrap.pack(fill="both", expand=True, pady=8)

        self.tree = ttk.Treeview(table_wrap, show="headings")
        self.tree.pack(side="left", fill="both", expand=True)

        yscroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        yscroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=yscroll.set)

        xscroll = ttk.Scrollbar(self.frame, orient="horizontal", command=self.tree.xview)
        xscroll.pack(fill="x")
        self.tree.configure(xscrollcommand=xscroll.set)

    # ---------- File IO ----------
    def load_file(self):
        path = filedialog.askopenfilename(
            title="Select Petty Cash Excel",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not path:
            return

        try:
            # Your template rule: header row = row 4 (0-based header=3)
            # and then drop the next 2 rows (which are rows 5 & 6 in Excel terms)
            df = pd.read_excel(path, header=3)
            df.columns = [str(c).strip() for c in df.columns]
            df = df.iloc[2:].reset_index(drop=True)

            self.src_df = df

            # Apply shared transform (petty/eWallet both use same rules)
            self.proc_df = transform_petty_or_ewallet(df)

            self._display(self.proc_df)
            messagebox.showinfo("Loaded", "File loaded and processed successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load/process file:\n{e}")

    def export_file(self, kind: str = "xlsx"):
        if self.proc_df is None or self.proc_df.empty:
            messagebox.showwarning("No Data", "Load a file first.")
            return

        if kind == "xlsx":
            path = filedialog.asksaveasfilename(
                title="Save Processed Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")]
            )
            if not path:
                return
            try:
                self.proc_df.to_excel(path, index=False)
                messagebox.showinfo("Exported", f"Saved to {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save Excel:\n{e}")

        elif kind == "csv":
            path = filedialog.asksaveasfilename(
                title="Save Processed CSV",
                defaultextension=".csv",
                filetypes=[("CSV", "*.csv")]
            )
            if not path:
                return
            try:
                self.proc_df.to_csv(path, index=False, encoding="utf-8-sig")
                messagebox.showinfo("Exported", f"Saved to {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save CSV:\n{e}")

    # ---------- Table display ----------
    def _display(self, df: pd.DataFrame):
        # Clear
        self.tree.delete(*self.tree.get_children())
        for col in self.tree["columns"]:
            self.tree.heading(col, text="")

        cols = list(df.columns)
        self.tree["columns"] = cols

        # Headings + simple width heuristic
        for col in cols:
            self.tree.heading(col, text=col)
            max_len = max([len(str(col))] + [len(str(v)) for v in df[col].head(200).tolist()])
            width = min(max(80, max_len * 8), 380)
            self.tree.column(col, width=width, anchor="w")

        # Rows
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=[row.get(c, "") for c in cols])
