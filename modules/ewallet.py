import pandas as pd

pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 0)
pd.set_option('display.max_colwidth', None)


# 1) Read with header on row 4 (0-based index=3) then drop rows 5 & 6
df = pd.read_excel("E-Wallet Template.xlsx", header=3, engine="openpyxl")
df.columns = df.columns.str.strip()
df = df.iloc[2:].reset_index(drop=True)
print(df)

# Case-insensitive column lookup
cols = {c.lower(): c for c in df.columns}
def col(name): return cols.get(name.lower())

# 2) Date filter (keep only real dates)
date_col = col("Date")
df[date_col] = pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")
df = df[df[date_col].notna()].reset_index(drop=True)

# 3) Build output
output_columns = ['TxDate','Description','Reference','Amount','UseTax','TaxType','TaxAccount','TaxAmount','Project','Account','IsDebit','SplitType','SplitGroup','Reconcile','PostDated','UseDiscount','DiscPerc','DiscTrCode','DiscDesc','UseDiscTax','DiscTaxType','DiscTaxAcc','DiscTaxAmt','PayeeName','PrintCheque','SalesRep','Module','SagePayExtra1','SagePayExtra2','SagePayExtra3']
defaults = {'TaxType':'','TaxAccount':'','TaxAmount':0,'Project':'','IsDebit':'N','SplitType':0,'SplitGroup':0,'Reconcile':'N','PostDated':'N','UseDiscount':'N','DiscPerc':0,'DiscTrCode':'','DiscDesc':'','UseDiscTax':'N','DiscTaxType':'','DiscTaxAcc':'','DiscTaxAmt':0,'PayeeName':'','PrintCheque':'N','SalesRep':'','Module':0,'SagePayExtra1':'','SagePayExtra2':'','SagePayExtra3':''}



# Format the DataFrame
formatted_df = pd.DataFrame(columns=output_columns)

# case-insensitive source columns
date_c = col('Date')
desc_c = col('Description')
ref_c  = col('Reference')
acct_c = col('Pastel_Acc')
paid_c = col('Amount_paid') or col('amount_paid')
recv_c = col('Amount_received') or col('amount_received')
vat_c  = col('Vat(Y/N)') or col('VAT')

# core fields
formatted_df['TxDate']      = df[date_c]
formatted_df['Description'] = (df[desc_c].astype(str) if desc_c else "")
formatted_df['Reference']   = (df[ref_c] if ref_c else "")
formatted_df['Account']     = (df[acct_c].astype(str).str.strip() if acct_c else "")

# amounts
idx  = df.index
paid = pd.to_numeric(df[paid_c], errors='coerce') if paid_c else pd.Series([pd.NA]*len(idx), index=idx, dtype='float64')
recv = pd.to_numeric(df[recv_c], errors='coerce') if recv_c else pd.Series([pd.NA]*len(idx), index=idx, dtype='float64')
formatted_df['Amount'] = paid.where(paid.notna() & paid.ne(0), recv).fillna(0).round(2)

# VAT -> UseTax
if vat_c:
    vat_series = df[vat_c].astype(str).str.strip().str.upper().str[0]
else:
    vat_series = pd.Series(['N'] * len(idx), index=idx)
formatted_df['UseTax'] = vat_series.map(lambda v: 'Y' if v == 'Y' else 'N')

# date format dd/mm/yyyy
formatted_df['TxDate'] = pd.to_datetime(formatted_df['TxDate'], dayfirst=True, errors='coerce').dt.strftime('%d/%m/%Y')

# fill defaults except Module/IsDebit (set below)
for k, v in defaults.items():
    if k not in ('Module', 'IsDebit'):
        formatted_df[k] = v

# Module mapping (gl->0, ap->1, ar->2)
mod_col = next((c for c in df.columns if c.lower() in ('mymodule','module')), None)
if mod_col:
    mod_map = {'gl': 0, 'ap': 2, 'ar': 1}
    formatted_df['Module'] = df[mod_col].astype(str).str.strip().str.lower().map(mod_map).fillna(0).astype(int)
else:
    formatted_df['Module'] = 0

# IsDebit: Y if value taken from Amount_received
mask_received = recv.notna() & recv.ne(0) & (paid.isna() | paid.eq(0))
formatted_df['IsDebit'] = mask_received.map({True: 'Y', False: 'N'})

# fill missing Reference with 'DEP'
ref_mask = formatted_df['Reference'].isna() | (formatted_df['Reference'].astype(str).str.strip() == '')
formatted_df.loc[ref_mask, 'Reference'] = 'DEP'

# final order
formatted_df = formatted_df[output_columns]



formatted_df.to_excel("petty_import.xlsx", index=False, engine="openpyxl")
formatted_df.to_csv("petty_import.csv", index=False, encoding="utf-8-sig")


