import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry  # this works fine without ttkbootstrap
import pandas as pd
import re
import os

class GLExtractorApp:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent, padding=20)
        self.frame.pack(fill="both", expand=True)

        self.df = None
        self.build_ui()

    def build_ui(self):
        ttk.Label(self.frame, text="📤 BFN Central - Recovery Invoices", font=("Segoe UI", 14, "bold")).pack(pady=(0, 10))

        # Instruction box
        info_frame = ttk.LabelFrame(self.frame, text="How it works", padding=10)
        info_frame.pack(fill="x", pady=10)

        steps = [
            "1. Select the last day of the month for transaction processing.",
            "2. Fetch the report from Reports > General Ledger > Account Transactions in Excel format. (Display=Print Account)",
            "3. Select the Excel file containing the ledger data (.xlsx format).",
            "4. The program will extract and process the data.",
            "5. The processed data will be saved as CSV files.",
            "6. The files will be named 'Inv_InvoiceCustomer.csv' and 'Inv_InvoiceSupplier.csv'.",
            "7. The customer invoice must be re-saved as a UTF-8 encoded file.",
            "8. The files will be saved in the same folder as the selected Excel file.",
            "9. The program will notify you once the files are saved.",
            "10. The program will handle credit-only rows by setting the quantity to -1.",
            "11. The program will only process rows with GLs ending in '/007'.",
        ]

        for step in steps:
            ttk.Label(info_frame, text=step, font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=2)

        # File selection area
        form = ttk.Frame(self.frame)
        form.pack(pady=(10, 15))

        ttk.Label(form, text="Select processing date").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.date_entry = DateEntry(form, width=20, date_pattern="dd/mm/yyyy")
        self.date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(form, text="Select a ledger Excel file").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        ttk.Button(form, text="📁 Browse File", command=self.process_file).grid(row=1, column=1, padx=5, pady=5)

        # Output preview
        self.filename_label = ttk.Label(self.frame, text="", foreground="blue")
        self.filename_label.pack(pady=5)

    def process_file(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
            if not file_path:
                return

            df = pd.read_excel(file_path, sheet_name=0, header=None, engine='openpyxl')
            cleaned_data = []
            current_gl = None

            for index, row in df.iterrows():
                first_cell = str(row[0]) if not pd.isna(row[0]) else ""
                if re.match(r"\d{4}/[A-Z]{2,3}/\d{3}", first_cell):
                    current_gl = first_cell.strip()
                    continue
                try:
                    date = pd.to_datetime(first_cell, errors='coerce')
                    if pd.notna(date):
                        cleaned_data.append([current_gl] + row.tolist())
                except:
                    continue

            columns = ['GL', 'Date', 'Reference', 'Description', 'Unused', 'Debit', 'Credit', 'Balance']
            clean_df = pd.DataFrame(cleaned_data, columns=columns)
            selected_date = self.date_entry.get()

            clean_df['DOCTYPE'] = 4
            clean_df['ACCOUNTID'] = 'A008'
            clean_df['DESCRIPTION'] = 'Aucor Central'
            clean_df['INVDATE'] = selected_date
            clean_df['TAXINCLUSIVE'] = ''
            clean_df['ORDERNUM'] = 'A008'
            clean_df['CDESCRIPTION'] = 'REC: ' + clean_df['Description'].astype(str)
            clean_df['CLINENOTES'] = ''
            clean_df['FQUANTITY'] = 1
            clean_df['FUNITPRICEEXCL'] = clean_df['Debit'].fillna(0).astype(float)

            credit_only_mask = (clean_df['Debit'].fillna(0) == 0) & (clean_df['Credit'].fillna(0) > 0)
            clean_df.loc[credit_only_mask, 'FUNITPRICEEXCL'] = clean_df.loc[credit_only_mask, 'Credit'].astype(float)
            clean_df.loc[credit_only_mask, 'FQUANTITY'] = -1
            clean_df['FQTYTOPROCESS'] = clean_df['FQUANTITY']

            clean_df['IMODULE'] = 1
            clean_df['ISTOCKCODEID'] = ''
            clean_df['ILEDGERACCOUNTID'] = clean_df['GL']
            clean_df['ITAXTYPEID'] = 1
            clean_df['IWAREHOUSEID'] = 'MSTR'
            clean_df['IPRICELISTNAMEID'] = 1

            clean_df = clean_df[clean_df['ILEDGERACCOUNTID'].str.endswith('/007')]

            final_df = clean_df[['DOCTYPE', 'ACCOUNTID', 'DESCRIPTION', 'INVDATE', 'TAXINCLUSIVE', 'ORDERNUM',
                                 'CDESCRIPTION', 'CLINENOTES', 'FQUANTITY', 'FQTYTOPROCESS', 'FUNITPRICEEXCL',
                                 'IMODULE', 'ISTOCKCODEID', 'ILEDGERACCOUNTID', 'ITAXTYPEID', 'IWAREHOUSEID',
                                 'IPRICELISTNAMEID']]

            final_df = final_df.replace({'\'': '', '\"': '', ',': ''}, regex=True)

            output_dir = os.path.dirname(file_path)
            cust_path = os.path.join(output_dir, "Inv_InvoiceCustomer.csv")
            final_df.to_csv(cust_path, index=False)

            supplier_df = final_df.copy()
            supplier_df['DOCTYPE'] = 5
            supplier_df['ACCOUNTID'] = 'A001'
            supplier_df['DESCRIPTION'] = 'Aucor Bloemfontein'
            supplier_df['ORDERNUM'] = 'A001'
            supplier_df['CDESCRIPTION'] = supplier_df['CDESCRIPTION'].str.replace('^REC:', 'B:', regex=True)
            supplier_df = supplier_df.replace({'\'': '', '\"': '', ',': ''}, regex=True)

            supp_path = os.path.join(output_dir, "Inv_InvoiceSupplier.csv")
            supplier_df.to_csv(supp_path, index=False)

            messagebox.showinfo("Success", f"Files saved:\n{cust_path}\n{supp_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n{str(e)}")
