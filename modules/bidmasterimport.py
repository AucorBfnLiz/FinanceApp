import re
import os
import pandas as pd
from tkinter import Label, IntVar, Radiobutton, Button, Entry, messagebox, Frame, W, E
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from datetime import datetime
from tkcalendar import DateEntry

BLOEMFONTEIN = 1
WITBANK = 2

DEPARTMENT_CODES = {
    'Bfn Mining': '005',
    'Bfn Warehouse': '007',
    'Bfn Vehicles': '015',
    'Witbank Mining': '005',
    'Witbank Vehicles': '015',
    'Bfn Gov & Other': '010',
    'Witbank Gov & Other': '010'
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
        self.radio_state = IntVar()
        self.build_ui()

    def build_ui(self):
        Label(self.parent, text='Bidmaster Sales Jnl', font=('arial', 30), fg='red').grid(column=0, row=0, columnspan=3, pady=10)
        

        radio_frame = Frame(self.parent)
        radio_frame.grid(column=0, row=1, padx=10, pady=10, sticky=W)
        Radiobutton(radio_frame, text="Bloemfontein", value=BLOEMFONTEIN, variable=self.radio_state).pack(anchor=W)
        Radiobutton(radio_frame, text="Witbank", value=WITBANK, variable=self.radio_state).pack(anchor=W)

        Label(self.parent, text="Auction Code: (000) without the 'B' or 'W'").grid(column=0, row=3, sticky=E, padx=5, pady=5)
        self.auction_code_entry = Entry(self.parent, width=20)
        self.auction_code_entry.grid(column=1, row=3, sticky=W)

        Label(self.parent, text="Auction Dept:").grid(column=0, row=5, sticky=E, padx=5, pady=5)
        self.dept_combobox = ttk.Combobox(self.parent, values=list(DEPARTMENT_CODES.keys()), width=20)
        self.dept_combobox.grid(column=1, row=5, sticky=W)

        Label(self.parent, text="Auction Commission:").grid(column=0, row=6, sticky=E, padx=5, pady=5)
        self.commission_code_entry = Entry(self.parent, width=20)
        self.commission_code_entry.grid(column=1, row=6, sticky=W)

        Label(self.parent, text="Select Auction Date:").grid(column=0, row=7, sticky=E, padx=5, pady=5)
        self.date_code_entry = DateEntry(self.parent, width=20, date_pattern='dd/mm/yyyy')
        self.date_code_entry.grid(column=1, row=7, sticky=W)

        Button(self.parent, text='Proceed', command=self.button_proceed_function).grid(column=1, row=8, sticky=W, padx=5, pady=5)

        Label(self.parent, text="You need to export the Detailed Profit Report and Cash Recon report to CSV").grid(column=0, row=9, sticky=E, padx=5, pady=5)

    def validate_inputs(self):
        if not self.radio_state.get():
            messagebox.showerror("Error", "Please select a location.")
            return False
        if not self.auction_code_entry.get():
            messagebox.showerror("Error", "Enter auction code.")
            return False
        if not self.commission_code_entry.get():
            messagebox.showerror("Error", "Enter commission %.")
            return False
        if not self.dept_combobox.get():
            messagebox.showerror("Error", "Select a department.")
            return False
        if not self.date_code_entry.get():
            messagebox.showerror("Error", "Select a date.")
            return False
        try:
            datetime.strptime(self.date_code_entry.get(), "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", "Invalid date format.")
            return False
        self.chosen_location = self.radio_state.get()
        self.chosen_auction_code = self.auction_code_entry.get()
        self.chosen_department = self.dept_combobox.get()
        self.chosen_department_code = DEPARTMENT_CODES[self.chosen_department]
        self.date = self.date_code_entry.get()
        self.chosen_commission = self.commission_code_entry.get()
        return True

    def detail_profit_report(self):
        try:
            messagebox.showinfo("Select File", "Select the Detailed Profit Report (CSV).")
            path = askopenfilename(filetypes=[('CSV Files', '*.csv')])
            if not path:
                return None
            df = pd.read_csv(path, header=None, encoding='ISO-8859-1', engine='python')
            df = df.iloc[:, :36]
            headers = [chr(i) for i in range(65, 91)] + [f"A{chr(i)}" for i in range(65, 75)]
            df.columns = headers
            return df
        except Exception as e:
            messagebox.showerror("Error", f"Load error: {e}")
            return None

    def cashrecon_report(self):
        try:
            path = askopenfilename(filetypes=[('CSV Files', '*.csv')])
            if not path:
                return None
            df = pd.read_csv(path, header=None, encoding='ISO-8859-1', engine='python')
            df = df.iloc[:, :26]
            df.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
                          'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'buyer_nr',
                          'aDescription', 'V', 'W', 'X', 'Y', 'Z']
            df['buyer_nr'] = df['buyer_nr'].astype(str).str.replace(':', '').str.strip()
            df['aDescription'] = df['aDescription'].astype(str).str.strip().str.title()
            return df
        except Exception as e:
            messagebox.showerror("Error", f"Cash recon error: {e}")
            return None

    def convert_file(self, invoice_df, extracted_df):
        branch_letter = 'B' if self.chosen_location == BLOEMFONTEIN else 'W'
        branch_gl = 'BLM' if self.chosen_location == BLOEMFONTEIN else 'WB'
        invoice_df.columns = invoice_df.columns.str.strip()
        if 'AF' not in invoice_df.columns:
            messagebox.showerror("Error", "Column 'AF' not found.")
            return None
        invoice_df['AF'] = invoice_df['AF'].astype(str).str.replace(' ', '').str.replace(',', '').replace('', '0')
        invoice_df['fUnitPriceExcl'] = invoice_df['AF'].astype(float)
        invoice_df['AB'] = invoice_df['AB'].astype(str).str.strip()
        extracted_df['buyer_nr'] = extracted_df['buyer_nr'].astype(str).str.strip()
        merged_df = pd.merge(invoice_df, extracted_df[['buyer_nr', 'aDescription']], left_on='AB', right_on='buyer_nr', how='left')
        invoice_df['aDescription'] = merged_df['aDescription']
        invoice_df['DocType'] = '4'
        invoice_df['AccountID'] = branch_letter + self.auction_code_entry.get() + '/P' + invoice_df['AB']
        invoice_df['InvDate'] = self.date
        invoice_df['TaxInclusive'] = ''
        invoice_df['OrderNum'] = invoice_df['AccountID']
        invoice_df['cDescription'] = 'Lot nr ' + invoice_df['AA'].astype(str) + ' - ' + invoice_df['M'].astype(str)


        invoice_df['CLIENTNOTES'] = invoice_df['AC']
        invoice_df['fQuantity'] = invoice_df['fQtyToProcess'] = '1'
        invoice_df['iModule'] = '1'
        invoice_df['iStockCodeID'] = ''
        invoice_df['iLedgerAccountID'] = f"8010/{branch_gl}/{self.chosen_department_code}/{self.chosen_auction_code}"
        invoice_df['iTaxTypeID'] = '20'
        invoice_df['iWarehouseID'] = 'MSTR'
        invoice_df['iPriceListNameID'] = '1'
        final_df = invoice_df[['DocType', 'AccountID', 'aDescription', 'InvDate', 'TaxInclusive', 'OrderNum',
                               'cDescription', 'CLIENTNOTES', 'fQuantity', 'fQtyToProcess', 'fUnitPriceExcl',
                               'iModule', 'iStockCodeID', 'iLedgerAccountID', 'iTaxTypeID', 'iWarehouseID', 'iPriceListNameID']]
        final_df.to_csv("Bidprice.csv", index=False)
        return final_df

    def add_commission(self, branch_gl, dept_code):
        try:
            df = pd.read_csv("Bidprice.csv")
            commission = float(self.chosen_commission)
            df['fUnitPriceExcl'] *= (commission / 100)
            df['iLedgerAccountID'] = f'1630/{branch_gl}/{dept_code}'
            df['iTaxTypeID'] = '1'
            df['CLIENTNOTES'] = ''
            df['cDescription'] = df['cDescription'].str.replace(r' - .*$', f' - Buyers Commission @ {commission}%', regex=True)
            return df
        except Exception as e:
            messagebox.showerror("Error", f"Commission error: {e}")
            return None

    def add_docfee(self, df, branch_gl, dept_code):
        docfee = df.copy()
        docfee['fUnitPriceExcl'] = 2600
        docfee['iLedgerAccountID'] = f'1990/{branch_gl}/{dept_code}'
        docfee['cDescription'] = docfee['cDescription'].str.replace(r'\s*-\s*.*$', ' - Documentation Fee', regex=True)
        return docfee

    def merge_dataframes(self, df1, df2, df3):
        merged = pd.concat([df1, df2, df3], ignore_index=True)
        merged['CLIENTNOTES'] = merged['CLIENTNOTES'].str.replace(r"[\"',]", '', regex=True)
        merged.to_csv('Inv_Invoice Import.csv', index=False)
        messagebox.showinfo("Done", "Invoice Import file created.")

    def button_proceed_function(self):
        if not self.validate_inputs():
            return
        branch_gl = 'BLM' if self.chosen_location == BLOEMFONTEIN else 'WB'
        if not messagebox.askyesno("Confirm", f"Proceed with {self.chosen_department} @ {self.chosen_commission}%?"):
            return
        invoice_df = self.detail_profit_report()
        if invoice_df is None:
            return
        cash_df = self.cashrecon_report()
        if cash_df is None:
            return
        final_df = self.convert_file(invoice_df, cash_df)
        if final_df is None:
            return
        final_comm = self.add_commission(branch_gl, self.chosen_department_code)
        if final_comm is None:
            return
        final_docfee = self.add_docfee(final_comm, branch_gl, self.chosen_department_code)
        if final_docfee is None:
            return
        self.merge_dataframes(final_df, final_comm, final_docfee)
