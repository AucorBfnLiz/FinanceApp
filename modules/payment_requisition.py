import pandas as pd
from datetime import datetime
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from tkinter import filedialog, messagebox, simpledialog
import tkinter as tk
from tkinter import ttk

class PDF(FPDF):
    def __init__(self, user):
        super().__init__()
        self.user = user

    def header(self):
        self.image('Aucor-Logo.png', 12, 1, 40)
        self.set_font('helvetica', 'B', 25)
        self.cell(0, 20, 'Aucor Payment Requisition', border=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
        self.ln(4)

    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 10)
        self.cell(0, 10, f'Payment requisition created by {self.user}', align='C')

    def chapter_title(self, title):
        self.set_font('helvetica', 'B', 12)
        self.cell(0, 10, title, 0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='L')
        self.ln(1)
        self.cell(0, 0, '', 'T', new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.ln(5)

    def add_form_field(self, label):
        self.set_font('helvetica', '', 10)
        self.cell(0, 8, label, 0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='L')
        self.ln(2)

    def add_approval_section(self):
        self.set_font('helvetica', 'B', 10)
        self.cell(0, 10, 'Approval Signatures:', 0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='L')
        self.cell(0, 0, '', 'T', new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
        self.add_form_field('Approver 1: Jacques van der Linde - Managing Director')
        self.add_form_field('Approver 2: Charles Neser - National Mining Manager')
        self.add_form_field('Signatures will be on Payment Requisition or Daily Transfer List')


class PaymentRequisitionApp:
    def __init__(self, parent):
        self.parent = parent
        self.frame = ttk.Frame(parent, padding=20)
        self.frame.pack(fill="both", expand=True)
        self.df = None
        self.user = None
        self.build_ui()

    def build_ui(self):
        ttk.Label(self.frame, text="📄 Payment Requisition Creator", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 15))

        info_frame = ttk.LabelFrame(self.frame, text="Instructions", padding=10)
        info_frame.pack(fill="x", pady=(0, 15))

        instructions = [
            "1. Click 'Load Customer Excel' to import your refund data (includes Group, Name, Customer Code, Balance, etc).",
            "2. Click 'Generate PDFs' to create payment requisition documents for customers.",
            "3. Use the Supplier section below for supplier payments (requires Supplier Code, Name, Bank info, etc).",
            "4. All balances will be auto-cleaned and formatted (R, commas removed).",
            "5. Output PDFs will be saved in the current working folder."
        ]
        for line in instructions:
            ttk.Label(info_frame, text=line, font=("Segoe UI", 10)).pack(anchor="w", padx=5, pady=2)

        ttk.Label(self.frame, text="📄 Payment Requisition for Customer Refunds", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 15))

        btn_frame = ttk.Frame(self.frame)
        btn_frame.pack(pady=10)

        ttk.Button(btn_frame, text="📥 Load Customer Excel", command=self.load_excel).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="📤 Generate PDFs", command=self.generate_pdfs).pack(side="left", padx=10)

        self.status_label = ttk.Label(self.frame, text="No file loaded yet.", foreground="blue")

        ttk.Separator(self.frame, orient="horizontal").pack(fill="x", pady=20)

        ttk.Label(self.frame, text="🏢 Payment Requisition for Suppliers", font=("Segoe UI", 14, "bold")).pack(anchor="w", pady=(0, 15))

        supplier_btn_frame = ttk.Frame(self.frame)
        supplier_btn_frame.pack(pady=10)

        ttk.Button(supplier_btn_frame, text="📥 Load Supplier Excel", command=self.load_supplier_excel).pack(side="left", padx=10)
        ttk.Button(supplier_btn_frame, text="📤 Generate Supplier PDFs", command=self.generate_supplier_pdfs).pack(side="left", padx=10)

        self.supplier_status_label = ttk.Label(self.frame, text="No supplier file loaded yet.", foreground="blue")
        self.supplier_status_label.pack(anchor="w", pady=5)
        self.status_label.pack(anchor="w", pady=5)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            self.df = pd.read_excel(path)
            self.df['Balance'] = (
                self.df['Balance']
                .astype(str)
                .str.replace(r'[Rr]', '', regex=True)
                .str.replace(',', '', regex=False)
                .str.strip()
                .astype(float)
                .round(2) * -1
            )
            self.df['Telephone 1'] = self.df['Telephone 1'].astype(str).str.replace('.0', '', regex=False)
            self.status_label.config(text=f"Loaded: {path}", foreground="green")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file:\n{e}")

    def generate_pdfs(self):
        if self.df is None:
            messagebox.showwarning("No Data", "Please load a customer Excel file first.")
            return

        self.user = simpledialog.askstring("User", "Enter your Name and Surname:")
        if not self.user:
            return

        for _, row in self.df.iterrows():
            self.create_pdf(row)

        messagebox.showinfo("Done", "All PDFs created successfully.")

    def load_supplier_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            self.supplier_df = pd.read_excel(path)
            self.supplier_df['Balance'] = (
                self.supplier_df['Balance']
                .astype(str)
                .str.replace(r'[Rr]', '', regex=True)
                .str.replace(',', '', regex=False)
                .str.strip()
                .astype(float)
                .round(2)
            )
            self.supplier_df[['Bank Name', 'Bank Branch Code', 'Bank Account No']] = (
                self.supplier_df[['Bank Name', 'Bank Branch Code', 'Bank Account No']].fillna('')
            )
            self.supplier_status_label.config(text=f"Loaded: {path}", foreground="green")
        except Exception as e:
            messagebox.showerror("Error", f"Could not load supplier file:\n{e}")

    def generate_supplier_pdfs(self):
        if not hasattr(self, 'supplier_df') or self.supplier_df is None:
            messagebox.showwarning("No Data", "Please load a supplier Excel file first.")
            return

        self.user = simpledialog.askstring("User", "Enter your Name and Surname:")
        if not self.user:
            return

        for _, row in self.supplier_df.iterrows():
            self.create_supplier_pdf(row)

        messagebox.showinfo("Done", "All supplier PDFs created successfully.")

    def create_pdf(self, row):
        pdf = PDF(self.user)
        pdf.add_page()

        auction_code = str(row['Group'])
        auction_name = str(row['Group Description'])
        client_name = str(row['Name'])
        client_company = str(row['Customer Description']) if pd.notna(row['Customer Description']) else ''
        client_customer_code = str(row['Customer']).replace('/', '_')
        client_phone = str(row['Telephone 1']) if pd.notna(row['Telephone 1']) else ''
        client_email = str(row['E-mail']) if pd.notna(row['E-mail']) else ''
        refund_amount = '{:,.2f}'.format(row['Balance'])

        today = datetime.today().strftime('%d/%m/%Y')
        pdf.set_font('helvetica', '', 12)
        pdf.cell(0, 8, f'Date: {today}', align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.set_font('helvetica', 'B', 12)
        pdf.cell(100, 10, 'Auction Refund/Deposit Refund', border=True)
        pdf.set_text_color(255, 0, 0)
        pdf.cell(90, 10, f'Amount to Refund: R {refund_amount}', align='R', border=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_text_color(0, 0, 0)
        pdf.ln(10)

        pdf.chapter_title('Auction Details:')
        pdf.add_form_field(f'Auction Code: {auction_code}')
        pdf.add_form_field(f'Auction Name: {auction_name}')

        pdf.chapter_title('Client Information')
        pdf.add_form_field(f'Customer Code: {client_customer_code}')
        pdf.add_form_field(f'Client Name: {client_name} : {client_company}')
        pdf.add_form_field(f'Phone Number: {client_phone}                      Email: {client_email}')

        pdf.chapter_title('Banking Details') 
        pdf.add_form_field('Bank Name:                                         Account Number:')
        pdf.add_form_field('ABSA: 632005      NED: 198765 ')
        pdf.add_form_field('STD:  051001      FNB: 250655 ')
        pdf.add_form_field('Capitec: 470010')
        pdf.add_form_field('Immediate Payment:                                 Mail Proof of payment:')

        pdf.chapter_title('Notes')
        pdf.add_form_field('')
        pdf.add_approval_section()

        file_name = f'{client_customer_code}_{client_name}.pdf'
        pdf.output(file_name)

    def create_supplier_pdf(self, row):
        pdf = PDF(self.user)
        pdf.add_page()

        auction_code = str(row['Group'])
        auction_name = str(row['Group Description'])
        client_name = str(row['Name'])
        client_customer_code = str(row['Supplier']).replace('/', '_')
        refund_amount = '{:,.2f}'.format(row['Balance'])
        bank_name = str(row.get('Bank Name', ''))
        bank_branch = str(row.get('Bank Branch Code', ''))
        bank_no = str(row.get('Bank Account No', ''))

        today = datetime.today().strftime('%d/%m/%Y')
        pdf.set_font('helvetica', '', 12)
        pdf.cell(0, 8, f'Date: {today}', align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.set_font('helvetica', 'B', 12)
        pdf.cell(100, 10, 'Supplier Payment', border=True)
        pdf.set_text_color(255, 0, 0)
        pdf.cell(90, 10, f'Amount to Refund: R {refund_amount}', align='R', border=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_text_color(0, 0, 0)
        pdf.ln(10)

        pdf.chapter_title('Auction Details:')
        pdf.add_form_field(f'Group: {auction_code}')
        pdf.add_form_field(f'Supplier Type: {auction_name}')

        pdf.chapter_title('Supplier Information')
        pdf.add_form_field(f'Supplier Code: {client_customer_code}')
        pdf.add_form_field(f'Supplier Name: {client_name} ')

        pdf.chapter_title('Banking Details') 
        pdf.add_form_field(f'Bank Name:  {bank_name}                                       Account Number:  {bank_no}')
        pdf.add_form_field(f'Bank Branch Code: {bank_branch}')
        pdf.add_form_field('(STD:  051001  FNB: 250655  ABSA: 632005  NED: 198765 Capitec: 470010)')
        pdf.add_form_field('Immediate Payment:                                 Mail Proof of payment:')

        pdf.chapter_title('Notes')
        pdf.add_form_field('')
        pdf.add_approval_section()

        file_name = f'{client_customer_code}_{client_name}.pdf'
        pdf.output(file_name)
