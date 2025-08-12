import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from tkcalendar import DateEntry

BLOEMFONTEIN = 1
WITBANK = 2

DEPARTMENT_CODES = {
    "Bfn Mining": "005",
    "Bfn Warehouse": "007",
    "Bfn Vehicles": "015",
    "Witbank Mining": "005",
    "Witbank Vehicles": "015",
    "Bfn Gov & Other": "010",
    "Witbank Gov & Other": "010",
}

class BidmasterSalesApp:
    def __init__(self, parent):
        self.parent = parent
        self.chosen_location = None
        self.chosen_auction_code = None
        self.chosen_department = None
        self.chosen_department_code = None
        self.chosen_commission = None
        self.date = None
        self.radio_state = tk.IntVar()

        self.build_ui()

    def build_ui(self):
        # Title
        ttk.Label(self.parent, text="📄 Bidmaster Sales Journal", font=("Segoe UI", 14, "bold")).pack(pady=(0, 10))

        # How it works
        info_frame = ttk.LabelFrame(self.parent, text="How it works", padding=10)
        info_frame.pack(fill="x", pady=10)

        steps = [
            "1. Select location (Bloemfontein or Witbank).",
            "2. Enter the auction code without the 'B' or 'W'.",
            "3. Choose the department from the dropdown.",
            "4. Enter the commission percentage.",
            "5. Select the auction date.",
            "6. Click Proceed to generate the Invoice Import CSV.",
            "7. The program will prompt for DPR and Cash Recon CSV files.",
            "8. It will create the base invoices, commission invoices, and doc fee invoices.",
            "9. The final file will be saved as 'Inv_Invoice Import.csv' in the program folder."
        ]
        for step in steps:
            ttk.Label(info_frame, text=step, font=("Segoe UI", 9)).pack(anchor="w", padx=10, pady=1)

        # Controls
        control_frame = ttk.Frame(self.parent)
        control_frame.pack(fill="x", pady=10)

        # Location radios
        location_frame = ttk.LabelFrame(control_frame, text="Location")
        location_frame.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        ttk.Radiobutton(location_frame, text="Bloemfontein", value=BLOEMFONTEIN, variable=self.radio_state).pack(anchor="w")
        ttk.Radiobutton(location_frame, text="Witbank", value=WITBANK, variable=self.radio_state).pack(anchor="w")

        # Auction code
        ttk.Label(control_frame, text="Auction Code:").grid(row=0, column=1, sticky="e", padx=5)
        self.auction_code_entry = ttk.Entry(control_frame, width=10)
        self.auction_code_entry.grid(row=0, column=2, sticky="w", padx=5)

        # Department
        ttk.Label(control_frame, text="Department:").grid(row=0, column=3, sticky="e", padx=5)
        self.dept_combobox = ttk.Combobox(control_frame, values=list(DEPARTMENT_CODES.keys()), width=20, state="readonly")
        self.dept_combobox.grid(row=0, column=4, sticky="w", padx=5)
        if self.dept_combobox["values"]:
            self.dept_combobox.current(0)

        # Commission
        ttk.Label(control_frame, text="Commission %:").grid(row=1, column=1, sticky="e", padx=5)
        self.commission_code_entry = ttk.Entry(control_frame, width=10)
        self.commission_code_entry.grid(row=1, column=2, sticky="w", padx=5)

        # Date
        ttk.Label(control_frame, text="Auction Date:").grid(row=1, column=3, sticky="e", padx=5)
        self.date_code_entry = DateEntry(control_frame, width=12, date_pattern="dd/mm/yyyy")
        self.date_code_entry.grid(row=1, column=4, sticky="w", padx=5)

        # Proceed button
        ttk.Button(self.parent, text="🚀 Proceed", command=self.button_proceed_function).pack(pady=10)

    def validate_inputs(self):
        if not self.radio_state.get():
            messagebox.showerror("Error", "Please select a location.")
            return False
        code = self.auction_code_entry.get().strip()
        if not code.isdigit():
            messagebox.showerror("Error", "Auction code must be numeric.")
            return False
        comm_text = self.commission_code_entry.get().strip().replace(",", ".")
        try:
            commission = float(comm_text)
        except ValueError:
            messagebox.showerror("Error", "Commission must be numeric.")
            return False
        if not self.dept_combobox.get():
            messagebox.showerror("Error", "Please select a department.")
            return False
        try:
            datetime.strptime(self.date_code_entry.get(), "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", "Invalid date.")
            return False

        self.chosen_location = self.radio_state.get()
        self.chosen_auction_code = code
        self.chosen_department = self.dept_combobox.get()
        self.chosen_department_code = DEPARTMENT_CODES[self.chosen_department]
        self.date = self.date_code_entry.get()
        self.chosen_commission = commission
        return True

    def detail_profit_report(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not path: return None
        df = pd.read_csv(path, header=None, encoding="ISO-8859-1", engine="python").iloc[:, :36]
        headers = [chr(i) for i in range(65, 91)] + [f"A{chr(i)}" for i in range(65, 75)]
        df.columns = headers
        return df

    def cashrecon_report(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not path: return None
        df = pd.read_csv(path, header=None, encoding="ISO-8859-1", engine="python").iloc[:, :26]
        df.columns = [
            "A","B","C","D","E","F","G","H","I","J",
            "K","L","M","N","O","P","Q","R","S","buyer_nr",
            "aDescription","V","W","X","Y","Z"
        ]
        df["buyer_nr"] = df["buyer_nr"].astype(str).str.replace(":", "", regex=False).str.strip()
        df["aDescription"] = df["aDescription"].astype(str).str.strip().str.title()
        return df

    def convert_file(self, invoice_df, extracted_df):
        branch_letter = "B" if self.chosen_location == BLOEMFONTEIN else "W"
        branch_gl = "BLM" if self.chosen_location == BLOEMFONTEIN else "WB"
        invoice_df["AF"] = invoice_df["AF"].astype(str).str.replace(" ", "", regex=False).str.replace(",", "", regex=False).replace("", "0")
        invoice_df["fUnitPriceExcl"] = invoice_df["AF"].astype(float)
        invoice_df["AB"] = invoice_df["AB"].astype(str).str.strip()
        merged_df = pd.merge(invoice_df, extracted_df[["buyer_nr", "aDescription"]], left_on="AB", right_on="buyer_nr", how="left")
        invoice_df["aDescription"] = merged_df["aDescription"].fillna("Buyer Unknown")
        invoice_df["DocType"] = "4"
        invoice_df["AccountID"] = branch_letter + self.chosen_auction_code + "/P" + invoice_df["AB"]
        invoice_df["InvDate"] = self.date
        invoice_df["TaxInclusive"] = ""
        invoice_df["OrderNum"] = invoice_df["AccountID"]
        invoice_df["cDescription"] = "Lot nr " + invoice_df["AA"].astype(str).str.strip() + " - " + invoice_df["M"].astype(str).str.strip()
        invoice_df["CLIENTNOTES"] = invoice_df["AC"]
        invoice_df["fQuantity"] = invoice_df["fQtyToProcess"] = "1"
        invoice_df["iModule"] = "1"
        invoice_df["iStockCodeID"] = ""
        invoice_df["iLedgerAccountID"] = f"8010/{branch_gl}/{self.chosen_department_code}"
        invoice_df["iTaxTypeID"] = "20"
        invoice_df["iWarehouseID"] = "MSTR"
        invoice_df["iPriceListNameID"] = "1"
        return invoice_df[[
            "DocType","AccountID","aDescription","InvDate","TaxInclusive","OrderNum",
            "cDescription","CLIENTNOTES","fQuantity","fQtyToProcess","fUnitPriceExcl",
            "iModule","iStockCodeID","iLedgerAccountID","iTaxTypeID","iWarehouseID","iPriceListNameID"
        ]]

    def add_commission(self, df, branch_gl, dept_code):
        out = df.copy()
        out["fUnitPriceExcl"] *= (self.chosen_commission / 100.0)
        out["iLedgerAccountID"] = f"1630/{branch_gl}/{dept_code}"
        out["iTaxTypeID"] = "1"
        out["CLIENTNOTES"] = ""
        out["cDescription"] = out["cDescription"].str.replace(r" - .*$", f" - Buyers Commission @ {self.chosen_commission}%", regex=True)
        return out

    def add_docfee(self, df, branch_gl, dept_code):
        out = df.copy()
        out["fUnitPriceExcl"] = 2600
        out["iLedgerAccountID"] = f"1990/{branch_gl}/{dept_code}"
        out["cDescription"] = out["cDescription"].str.replace(r"\s*-\s*.*$", " - Documentation Fee", regex=True)
        return out

    def merge_dataframes(self, *dfs):
        merged = pd.concat(dfs, ignore_index=True)
        merged["CLIENTNOTES"] = merged["CLIENTNOTES"].fillna("").astype(str).str.translate(str.maketrans({'"': '', "'": '', ',': '', '\\': ''}))
        merged.to_csv("Inv_Invoice Import.csv", index=False, encoding="utf-8-sig")
        messagebox.showinfo("Done", "Invoice Import file created: 'Inv_Invoice Import.csv'")

    def button_proceed_function(self):
        if not self.validate_inputs():
            return
        if not messagebox.askyesno("Confirm", f"Proceed with {self.chosen_department} @ {self.chosen_commission}%?"):
            return
        invoice_df = self.detail_profit_report()
        if invoice_df is None: return
        cash_df = self.cashrecon_report()
        if cash_df is None: return
        final_df = self.convert_file(invoice_df, cash_df)
        branch_gl = "BLM" if self.chosen_location == BLOEMFONTEIN else "WB"
        final_comm = self.add_commission(final_df, branch_gl, self.chosen_department_code)
        final_docfee = self.add_docfee(final_comm, branch_gl, self.chosen_department_code)
        self.merge_dataframes(final_df, final_comm, final_docfee)
 