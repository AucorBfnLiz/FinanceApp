import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import threading

class Compare9500App:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent, padding=20)
        self.frame.pack(fill="both", expand=True)

        self.df_a = None
        self.df_b = None

        self.build_ui()

    def build_ui(self):
        ttk.Label(self.frame, text="📊 Compare 9500 Reports", font=("Segoe UI", 14, "bold")).pack(pady=(0, 10))

        # Instructions
        info_frame = ttk.LabelFrame(self.frame, text="How it works", padding=10)
        info_frame.pack(fill="x", pady=10)

        steps = [
            "1. Select the number of columns you want to compare (usually 9).",
            "2. Upload Excel from Evolution).",
            "3. Upload Excel Reconciliation export).",
            "4. Click 'Compare' to find unmatched rows.",
            "5. Preview both unmatched sets (A and B) below. B should be empty except for totals. If not, check the files and redo the process.",
            "6. Export results of A and copy and paste it in to the reconciliation. Ensure the balance is correct.",
        ]
        for step in steps:
            ttk.Label(info_frame, text=step, font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=1)

        # Column selector
        control_frame = ttk.Frame(self.frame)
        control_frame.pack(pady=10)

        ttk.Label(control_frame, text="Number of columns to compare:").grid(row=0, column=0, padx=5)
        self.column_var = tk.IntVar(value=9)
        ttk.Spinbox(control_frame, from_=1, to=50, textvariable=self.column_var, width=5).grid(row=0, column=1, padx=5)

        # File buttons
        ttk.Button(control_frame, text="📁 Upload Evolution", command=self.load_file_a).grid(row=0, column=2, padx=10)
        ttk.Button(control_frame, text="📁 Upload Recon", command=self.load_file_b).grid(row=0, column=3, padx=10)
        ttk.Button(control_frame, text="🔍 Compare", command=self.compare).grid(row=0, column=4, padx=10)

        # Result label
        self.result_label = ttk.Label(self.frame, text="", font=("Segoe UI", 10, "bold"))
        self.result_label.pack(pady=(5, 10))

        # Treeviews for previews
        self.tree_frame = ttk.Frame(self.frame)
        self.tree_frame.pack(fill="both", expand=True)

        self.tree_frame.columnconfigure(0, weight=1)
        self.tree_frame.columnconfigure(1, weight=1)

        self.tree_a = self.create_treeview("Only in A", 0)
        self.tree_b = self.create_treeview("Only in B", 1)

        # Export buttons
        action_frame = ttk.Frame(self.frame)
        action_frame.pack(pady=10)

        ttk.Button(action_frame, text="⬇️ Export A - not on recon'", command=self.export_a).grid(row=0, column=0, padx=15)
        ttk.Button(action_frame, text="⬇️ Export B - twice on recon'", command=self.export_b).grid(row=0, column=1, padx=15)

    def create_treeview(self, title, col_index):
        container = ttk.LabelFrame(self.tree_frame, text=title, padding=5)
        container.grid(row=0, column=col_index, sticky="nsew", padx=5, pady=5)

        tree = ttk.Treeview(container, show="headings")
        tree.pack(side="left", fill="both", expand=True)

        scrollbar_y = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
        scrollbar_y.pack(side="right", fill="y")
        tree.configure(yscrollcommand=scrollbar_y.set)

        return tree

    def load_file_a(self):
        def task():
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if path:
                try:
                    df = pd.read_excel(path).iloc[:, :self.column_var.get()]
                    self.clean_dataframe(df)
                    self.df_a = df
                    self.parent.after(0, lambda: messagebox.showinfo("Success", "Excel A loaded successfully."))
                except Exception as e:
                    self.parent.after(0, lambda e=e: messagebox.showerror("Error", f"Failed to load Excel A:\n{e}"))

        threading.Thread(target=task).start()

    def load_file_b(self):
        def task():
            path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if path:
                try:
                    xls = pd.ExcelFile(path)
                    sheet = xls.sheet_names[0]
                    df = pd.read_excel(path, sheet_name=sheet).iloc[:, :self.column_var.get()]
                    self.clean_dataframe(df)
                    self.df_b = df
                    self.parent.after(0, lambda: messagebox.showinfo("Success", f"Excel B sheet '{sheet}' loaded."))
                except Exception as e:
                    self.parent.after(0, lambda e=e: messagebox.showerror("Error", f"Failed to load Excel B:\n{e}"))
        
        threading.Thread(target=task).start()

    def clean_dataframe(self, df):
        df.columns = df.columns.str.strip()

        for col in ["Debit", "Credit"]:
            if col in df.columns:
                df[col] = pd.to_numeric(
                    df[col].astype(str)
                    .str.replace("R", "", regex=False)
                    .str.replace(",", "", regex=False)
                    .str.strip(),
                    errors="coerce"
                ).fillna(0).round(2)

        for ref_col in ["Reference", "Reference 2"]:
            if ref_col in df.columns:
                df[ref_col] = df[ref_col].astype(str).str.strip().str.lstrip("0").str.replace(r"\.0$", "", regex=True)
                df[ref_col] = df[ref_col].replace(["", "None", "nan"], pd.NA)

        if "Date" in df.columns:
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce").dt.date

        if "Description" in df.columns:
            df["Description"] = df["Description"].astype(str).str.lstrip("0").str.strip()

    def compare(self):
        if self.df_a is None or self.df_b is None:
            messagebox.showwarning("Missing", "Please upload both Excel files first.")
            return

        if list(self.df_a.columns) != list(self.df_b.columns):
            messagebox.showerror(
                "Header Mismatch",
                "The column headers in Excel A and Excel B do not match.\n\n"
                "Please check that both files have the same structure before continuing."
            )
            return

        df_a_clean = self.df_a.dropna(how="all").fillna("")
        df_b_clean = self.df_b.dropna(how="all").fillna("")

        for col in df_a_clean.columns:
            if "date" in col.lower():
                df_a_clean[col] = pd.to_datetime(df_a_clean[col], errors="coerce").dt.date
                df_b_clean[col] = pd.to_datetime(df_b_clean[col], errors="coerce").dt.date
            elif col.lower() in ["debit", "credit"]:
                df_a_clean[col] = df_a_clean[col].astype(float)
                df_b_clean[col] = df_b_clean[col].astype(float)
            else:
                df_a_clean[col] = df_a_clean[col].astype(str)
                df_b_clean[col] = df_b_clean[col].astype(str)

        self.only_in_a = pd.merge(df_a_clean, df_b_clean, how="outer", indicator=True)\
                          .query("_merge == 'left_only'").drop(columns=["_merge"])
        self.only_in_b = pd.merge(df_b_clean, df_a_clean, how="outer", indicator=True)\
                          .query("_merge == 'left_only'").drop(columns=["_merge"])

        self.display_treeview(self.tree_a, self.only_in_a)
        self.display_treeview(self.tree_b, self.only_in_b)

        result = f"✅ Compared!\nOnly in A: {len(self.only_in_a)} rows\nOnly in B: {len(self.only_in_b)} rows"
        self.result_label.config(text=result)

    def display_treeview(self, tree, df):
        tree.delete(*tree.get_children())
        tree["columns"] = list(df.columns)

        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="center")

        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

    def export_a(self):
        if hasattr(self, "only_in_a") and not self.only_in_a.empty:
            path = filedialog.asksaveasfilename(defaultextension=".xlsx")
            if path:
                self.only_in_a.to_excel(path, index=False)
                messagebox.showinfo("Exported", f"Saved to {path}")

    def export_b(self):
        if hasattr(self, "only_in_b") and not self.only_in_b.empty:
            path = filedialog.asksaveasfilename(defaultextension=".xlsx")
            if path:
                self.only_in_b.to_excel(path, index=False)
                messagebox.showinfo("Exported", f"Saved to {path}")
