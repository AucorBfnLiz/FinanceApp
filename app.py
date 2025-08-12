import tkinter as tk
from tkinter import ttk

from modules.payment_requisition import PaymentRequisitionApp
from modules.import9500 import DepositImportApp
from modules.expenses import GLExtractorApp
from modules.bidmasterimport import BidmasterSalesApp
from modules.compare9500 import Compare9500App
from modules.petty_cash import PettyCashApp





root = tk.Tk()
root.geometry('900x650')
root.title('Finance Automation Toolkit')
# <-- must be before any ttk widgets



# Main notebook
notebook = ttk.Notebook(root)
notebook.pack(pady=10, expand=True, fill='both')

# Tab 1 - Welcome
welcome_frame = ttk.Frame(notebook)
ttk.Label(welcome_frame, text="Welcome to the Finance Toolkit", style="Title.TLabel").pack(pady=20)
notebook.add(welcome_frame, text='Welcome')

# Tab 2 - Petty Cash & eWallet (nested notebook)
central_rec = ttk.Notebook(notebook)
PettyCashApp(central_rec)
notebook.add(central_rec, text='EWallet_Petty Cash')

# Tab 3 - Compare 9500
compare9500 = ttk.Frame(notebook)
Compare9500App(compare9500)
notebook.add(compare9500, text='Compare 9500')

# Tab 4 - Import Deposits
import_deposits = ttk.Frame(notebook)
DepositImportApp(import_deposits)
notebook.add(import_deposits, text='Import Deposits')

# Tab 5 - Payment Requisitions
payment_reqs = ttk.Frame(notebook)
PaymentRequisitionApp(payment_reqs)
notebook.add(payment_reqs, text='💸 Requisitions')

# Tab 6 - Bidmaster Sales Jnl
bidmaster = ttk.Frame(notebook)
BidmasterSalesApp(bidmaster)
notebook.add(bidmaster, text='🧾 Bidmaster')

# Tab 7 - GL Extractor
gl_extract = ttk.Frame(notebook)
GLExtractorApp(gl_extract)
notebook.add(gl_extract, text='📂 GL Extractor')

root.mainloop()
