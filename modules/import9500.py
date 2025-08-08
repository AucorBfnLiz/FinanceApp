import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
from datetime import datetime

class DepositImportApp:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent, padding=20)
        self.frame.pack(fill="both", expand=True)
        self.df = None
        self.build_ui()

    def build_ui(self):
        ttk.Label(self.frame, text="🏦 Import Deposits from 9500 to Bank", font=("Segoe UI", 14, "bold")).pack(pady=(0, 15), anchor="w")

        info_frame = ttk.LabelFrame(self.frame, text="Instructions", padding=10)
        info_frame.pack(fill="x", pady=10)

        instructions = [
            "1. Upload an Excel file with columns: Date, Reference 2, Code, Reference, Description, Credit.",
            "2. The data will be formatted for Evolution.",
            "3. You'll be able to preview and export the processed file as CSV.",
        ]
        for line in instructions:
            ttk.Label(info_frame, text=line, font=("Segoe UI", 10)).pack(anchor="w", padx=5, pady=2)

        action_frame = ttk.Frame(self.frame)
        action_frame.pack(pady=15)

        ttk.Button(action_frame, text="📥 Upload Excel", command=self.load_excel).grid(row=0, column=0, padx=10)
        ttk.Button(action_frame, text="⬇️ Download CSV", command=self.download_csv).grid(row=0, column=1, padx=10)

        ttk.Label(self.frame, text="Preview:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10, 5))
        self.preview = tk.Text(self.frame, height=10, width=100, wrap="none", font=("Consolas", 9))
        self.preview.pack(fill="x", pady=5)

        self.filename_label = ttk.Label(self.frame, text="", font=("Segoe UI", 9), foreground="blue")
        self.filename_label.pack(pady=5, anchor="w")

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[
            ("Excel files", "*.xlsx *.xls *.xlsm"),
            ("All files", "*.*")
        ])
        if not path:
            return

        try:
            self.excel_path = path
            self.excel_file = pd.ExcelFile(path, engine="openpyxl")
            self.load_selected_sheet(self.excel_file.sheet_names[0])
        except Exception as e:
            messagebox.showerror("Error", f"Unable to open file:\n{e}")

    def load_selected_sheet(self, sheet):
        try:
            df = pd.read_excel(self.excel_file, sheet_name=sheet)
            df.columns = df.columns.str.strip()

            required = ["Date", "Description", "Credit", "Code", "Reference"]
            if any(col not in df.columns for col in required):
                messagebox.showerror("Missing Columns", f"File must contain: {', '.join(required)}")
                return

            df['Reference'] = df['Reference'].astype(str).str.strip()

            modified = pd.DataFrame()
            modified["TxDate"] = df["Date"]
            modified["Description"] = df["Description"]
            modified["Reference"] = df["Code"].astype(str) + df["Reference"].astype(str)
            modified["Amount"] = df["Credit"]
            modified["UseTax"] = "N"
            modified["TaxType"] = ""
            modified["TaxAccount"] = ""
            modified["TaxAmount"] = 0
            modified["Project"] = ""
            modified["Account"] = "9500/BLM/027"
            modified["IsDebit"] = "Y"
            modified["SplitType"] = 0
            modified["SplitGroup"] = 0
            modified["Reconcile"] = "N"
            modified["PostDated"] = "N"
            modified["UseDiscount"] = "N"
            modified["DiscPerc"] = 0
            modified["DiscTrCode"] = ""
            modified["DiscDesc"] = ""
            modified["UseDiscTax"] = "N"
            modified["DiscTaxType"] = ""
            modified["DiscTaxAcc"] = ""
            modified["DiscTaxAmt"] = 0
            modified["PayeeName"] = ""
            modified["PrintCheque"] = "N"
            modified["SalesRep"] = ""
            modified["Module"] = 0
            modified["SagePayExtra1"] = ""
            modified["SagePayExtra2"] = ""
            modified["SagePayExtra3"] = ""

            phrases = [
                "FNB APP PAYMENT FROM", "DIGITAL PAYMENT CR ABSA BANK", "CAPITEC",
                "ACB CREDIT CAPITEC", "FNB OB PMT", "PayShap Ext Credit",
                "INT-BANKING PMT FRM", "IMMEDIATE TRF CR CAPITEC",
                "IMMEDIATE TRF CR", "ACB CREDIT", "INVESTECPB"
            ]
            modified["Description"] = modified["Description"].replace(phrases, '', regex=True).str.strip()
            modified["Reference"] = modified["Reference"].str.strip().str.replace(r"\.0$", "", regex=True)

            self.df = modified
            self.preview.delete("1.0", "end")
            self.preview.insert("end", str(modified.head()))
            self.filename_label.config(text=f"Loaded file: {self.excel_path} (Sheet: {sheet})")

            messagebox.showinfo("Success", f"Excel sheet '{sheet}' processed successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process sheet:\n{e}")

    def download_csv(self):
        if self.df is None:
            messagebox.showwarning("Missing", "Please upload and process a file first.")
            return

        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                            initialfile=f"IMPORT_{datetime.today().strftime('%Y-%m-%d')}.csv")
        if path:
            self.df.to_csv(path, index=False)
            self.filename_label.config(text=f"Exported CSV to: {path}")
            messagebox.showinfo("Exported", f"CSV saved to:\n{path}")
