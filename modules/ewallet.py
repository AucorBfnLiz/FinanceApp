import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class EwalletApp:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent, padding=20)
        self.frame.pack(fill="both", expand=True)

        self.df = None
        self.build_ui()

    def build_ui(self):
        ttk.Label(self.frame, text="💳 eWallet Management", font=("Segoe UI", 14, "bold")).pack(pady=(0, 10))

        # Instruction box
        info_frame = ttk.LabelFrame(self.frame, text="How it works", padding=10)
        info_frame.pack(fill="x", pady=10)

        steps = [
            "1. Copy the eWallet file and delete all headers, subheaders, totals, etc.",
            "2. Run the Finance Toolbox and select the eWallet tab.",
            "3. Select the eWallet file you want to process.",
            "4. The program will process the file and generate a summary.",
        ]

        for step in steps:
            ttk.Label(info_frame, text=step, font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=2)

        # File selection area
        form = ttk.Frame(self.frame)
        form.pack(pady=(10, 15))

        ttk.Label(form, text="Select eWallet file").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Button(form, text="📁 Browse File", command=self.process_file).grid(row=0, column=1, padx=5, pady=5)

        # Output preview
        self.filename_label = ttk.Label(self.frame, text="", foreground="blue")
        self.filename_label.pack(pady=(10, 0))

    def process_file(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
            if not file_path:
                return

            df = pd.read_excel(file_path)
            if df.empty:
                messagebox.showerror("Error", "The selected file is empty or invalid.")
                return

            self.df = df
            self.filename_label.config(text=f"Processing file: {os.path.basename(file_path)}")
            self.process_ewallet_data()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while processing the file:\n{str(e)}")

    def process_ewallet_data(self):
        try:
            df = self.df
            if df is None or df.empty:
                messagebox.showerror("Error", "No data loaded.")
                return

            output_columns = [
                'TxDate', 'Description', 'Reference', 'Amount', 'UseTax', 'TaxType', 'TaxAccount', 'TaxAmount',
                'Project', 'Account', 'IsDebit', 'SplitType', 'SplitGroup', 'Reconcile', 'PostDated', 'UseDiscount',
                'DiscPerc', 'DiscTrCode', 'DiscDesc', 'UseDiscTax', 'DiscTaxType', 'DiscTaxAcc', 'DiscTaxAmt',
                'PayeeName', 'PrintCheque', 'SalesRep', 'Module', 'SagePayExtra1', 'SagePayExtra2', 'SagePayExtra3'
            ]

            formatted_df = pd.DataFrame(columns=output_columns)

            formatted_df['TxDate'] = df['Date']
            formatted_df['Description'] = df['Person_request'].astype(str) + " - " + df['Description'].astype(str)
            formatted_df['Reference'] = df['Reference']
            formatted_df['Amount'] = df['Amount_paid'].fillna(df['Amount_received'])
            formatted_df['UseTax'] = df['Vat(Y/N)']
            formatted_df['Account'] = df['Pastel_Acc']

            # Defaults
            defaults = {
                'TaxType': '', 'TaxAccount': '', 'TaxAmount': 0,
                'Project': '', 'IsDebit': 'N', 'SplitType': 0, 'SplitGroup': 0,
                'Reconcile': 'N', 'PostDated': 'N', 'UseDiscount': 'N', 'DiscPerc': 0,
                'DiscTrCode': '', 'DiscDesc': '', 'UseDiscTax': 'N', 'DiscTaxType': '',
                'DiscTaxAcc': '', 'DiscTaxAmt': 0, 'PayeeName': '', 'PrintCheque': 'N',
                'SalesRep': '', 'Module': 0, 'SagePayExtra1': '', 'SagePayExtra2': '', 'SagePayExtra3': ''
            }

            for col, val in defaults.items():
                formatted_df[col] = val

            formatted_df = formatted_df[output_columns]

            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save formatted eWallet file as"
            )

            if save_path:
                formatted_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"Formatted file saved:\n{save_path}")
            else:
                messagebox.showinfo("Cancelled", "Export cancelled.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during processing:\n{str(e)}")
