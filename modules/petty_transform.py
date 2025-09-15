# modules/petty_transform.py
import pandas as pd
from typing import Optional

OUTPUT_COLUMNS = [
    'TxDate','Description','Reference','Amount','UseTax','TaxType','TaxAccount','TaxAmount',
    'Project','Account','IsDebit','SplitType','SplitGroup','Reconcile','PostDated','UseDiscount',
    'DiscPerc','DiscTrCode','DiscDesc','UseDiscTax','DiscTaxType','DiscTaxAcc','DiscTaxAmt',
    'PayeeName','PrintCheque','SalesRep','Module','SagePayExtra1','SagePayExtra2','SagePayExtra3'
]

DEFAULTS = {
    'TaxType':'', 'TaxAccount':'', 'TaxAmount':0, 'Project':'', 'IsDebit':'N',
    'SplitType':0, 'SplitGroup':0, 'Reconcile':'N', 'PostDated':'N', 'UseDiscount':'N',
    'DiscPerc':0, 'DiscTrCode':'', 'DiscDesc':'', 'UseDiscTax':'N', 'DiscTaxType':'',
    'DiscTaxAcc':'', 'DiscTaxAmt':0, 'PayeeName':'', 'PrintCheque':'N', 'SalesRep':'',
    'Module':0, 'SagePayExtra1':'', 'SagePayExtra2':'', 'SagePayExtra3':''
}

def _lower_map(columns) -> dict:
    """Map lowercase stripped name -> original, tolerating NaN/None."""
    m = {}
    for c in columns:
        key = str(c).strip().lower()
        if key:  # ignore blanks
            m[key] = c
    return m

def _col(cols_map: dict, name: str) -> Optional[str]:
    return cols_map.get(name.strip().lower())

def transform_petty_or_ewallet(df: pd.DataFrame) -> pd.DataFrame:
    """
    Apply the shared petty/eWallet transformation rules to a DataFrame
    that has already been read with header row = 4 (0-based header=3)
    and the first two data rows after header removed.
    """
    # standardize column names
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    cols = _lower_map(df.columns)

    # --- Date filter (keep only real dates) ---
    date_c = _col(cols, 'date')
    if not date_c:
        raise ValueError("No 'Date' column found.")
    df[date_c] = pd.to_datetime(df[date_c], dayfirst=True, errors='coerce')
    df = df[df[date_c].notna()].reset_index(drop=True)

    # --- Source columns (case-insensitive) ---
    desc_c = _col(cols, 'description')
    ref_c  = _col(cols, 'reference')
    acct_c = _col(cols, 'pastel_acc') or _col(cols, 'pastel_Acc') or _col(cols, 'account')
    paid_c = _col(cols, 'amount_paid')
    recv_c = _col(cols, 'amount_received')
    vat_c  = _col(cols, 'vat(y/n)') or _col(cols, 'vat')

    # --- Build formatted output skeleton ---
    out = pd.DataFrame(index=df.index, columns=OUTPUT_COLUMNS)

    # core fields
    out['TxDate'] = df[date_c]
    out['Description'] = df[desc_c].astype(str) if desc_c else pd.Series('', index=df.index)
    out['Reference']   = df[ref_c] if ref_c else pd.Series('', index=df.index)
    out['Account']     = df[acct_c].astype(str).str.strip() if acct_c else pd.Series('', index=df.index)

    # amounts (prefer paid, else received)
    paid = pd.to_numeric(df[paid_c], errors='coerce') if paid_c else pd.Series(pd.NA, index=df.index, dtype='float64')
    recv = pd.to_numeric(df[recv_c], errors='coerce') if recv_c else pd.Series(pd.NA, index=df.index, dtype='float64')
    out['Amount'] = paid.where(paid.notna() & paid.ne(0), recv).fillna(0).round(2)

    # VAT -> UseTax
    if vat_c:
        vat_series = df[vat_c].astype(str).str.strip().str.upper().str[0]
    else:
        vat_series = pd.Series('N', index=df.index)
    out['UseTax'] = vat_series.eq('Y').map({True: 'Y', False: 'N'})

    # TxDate format dd/mm/yyyy (text)
    out['TxDate'] = pd.to_datetime(out['TxDate'], errors='coerce', dayfirst=True).dt.strftime('%d/%m/%Y')

    # Defaults (except Module / IsDebit set below)
    for k, v in DEFAULTS.items():
        if k not in ('Module', 'IsDebit'):
            out[k] = v

    # Module mapping: gl->0, ar->1, ap->2
    mod_col = next((c for c in df.columns if str(c).strip().lower() in ('mymodule','module')), None)
    if mod_col:
        mod_map = {'gl': 0, 'ar': 1, 'ap': 2}
        out['Module'] = (
            df[mod_col].astype(str).str.strip().str.lower().map(mod_map).fillna(0).astype(int)
        )
    else:
        out['Module'] = 0

    # IsDebit: Y if value came from Amount_received (and paid missing/0)
    mask_received = recv.notna() & recv.ne(0) & (paid.isna() | paid.eq(0))
    out['IsDebit'] = mask_received.map({True: 'Y', False: 'N'})

    # Reference: blank -> 'DEP'
    ref_mask = out['Reference'].isna() | (out['Reference'].astype(str).str.strip() == '')
    out.loc[ref_mask, 'Reference'] = 'DEP'

    # Final order & return
    out = out[OUTPUT_COLUMNS]
    return out

# --- Optional: run as a script on a file ---
if __name__ == "__main__":
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 0)
    pd.set_option('display.max_colwidth', None)

    # 1) Read with header on row 4 (index=3) then drop rows 5 & 6
    src_path = "E-Wallet Template.xlsx"  # or "Petty Cash Template.xlsx"
    df_in = pd.read_excel(src_path, header=3)      # engine autodetected
    df_in = df_in.iloc[2:].reset_index(drop=True)  # drop rows 5 & 6 (after header)

    formatted = transform_petty_or_ewallet(df_in)

    formatted.to_excel("petty_import.xlsx", index=False)
    formatted.to_csv("petty_import.csv", index=False, encoding="utf-8-sig")
    print("Saved petty_import.xlsx and petty_import.csv")
