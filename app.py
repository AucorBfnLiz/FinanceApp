import tkinter as tk
from tkinter import ttk
from modules.payment_requisition import PaymentRequisitionApp
from modules.import9500 import DepositImportApp
from modules.expenses import GLExtractorApp
from modules.bidmasterimport import BidmasterSalesApp
from modules.compare9500 import Compare9500App

# Root window
root = tk.Tk()
root.geometry('600x400')
root.title('Finance Automation Toolkit')

# Create custom style
style = ttk.Style()
style.configure("TNotebook.Tab", font=("Segoe UI", 12))

# Notebook (tabs)
notebook = ttk.Notebook(root)
notebook.pack(pady=10, expand=True, fill='both')

# Tab 1 - Welcome
welcome_frame = ttk.Frame(notebook)
welcome_frame.pack(fill='both', expand=True)
ttk.Label(welcome_frame, text="Welcome to the Finance Toolkit", font=("Segoe UI", 14)).pack(pady=20)
notebook.add(welcome_frame, text='Welcome')

# Tab 2 - Bfn-Cen Recover
central_rec = ttk.Frame(notebook)
central_rec.pack(fill='both', expand=True)
notebook.add(central_rec, text='Bfn-Cen Recover')

# Tab 3 - Compare 9500
compare9500 = ttk.Frame(notebook)
compare9500.pack(fill='both', expand=True)
Compare9500App(compare9500)
notebook.add(compare9500, text='Compare 9500')

# Tab 4 - Import Deposits
import_deposits = ttk.Frame(notebook)
import_deposits.pack(fill='both', expand=True)
DepositImportApp(import_deposits)
notebook.add(import_deposits, text='Import Deposits')

# Tab 5 - Payment Requisitions
payment_reqs = ttk.Frame(notebook)
payment_reqs.pack(fill='both', expand=True)
PaymentRequisitionApp(payment_reqs)
notebook.add(payment_reqs, text='💸 Requisitions')

# Tab 6 - Bidmaster Sales Jnl
bidmaster = ttk.Frame(notebook)
bidmaster.pack(fill='both', expand=True)
BidmasterSalesApp(bidmaster)
notebook.add(bidmaster, text='🧾 Bidmaster')

# Tab 7 - GL Extractor
gl_extract = ttk.Frame(notebook)
gl_extract.pack(fill='both', expand=True)
GLExtractorApp(gl_extract)
notebook.add(gl_extract, text='📂 GL Extractor')

# Start the app
root.mainloop()
