# test.py
import os
import pandas as pd

# --- 1) Pick your CSV ---
csv_path = "SMS AUG 2025.csv"
if not csv_path or not os.path.isfile(csv_path):
    raise FileNotFoundError(f"File not found: {csv_path}")

# --- 2) Read CSV with header on row 2 ---
df = pd.read_csv(csv_path, header=1)
df.columns = [str(c).strip() for c in df.columns]

print("\n=== Source headers ===")
print(list(df.columns))

# --- 3) Build inputs we need ---
msg_col = "Message Subject" if "Message Subject" in df.columns else None
date_col = "Send Date" if "Send Date" in df.columns else None
sms_col = "SMSs credit used" if "SMSs credit used" in df.columns else None

# Description
if msg_col:
    description = "Everlytic - " + df[msg_col].astype(str)
else:
    description = pd.Series([""] * len(df), index=df.index)

# TxDate formatted as DDMMYYYY
if date_col:
    txdate = pd.to_datetime(df[date_col], errors="coerce").dt.strftime("%d%m%Y")
else:
    txdate = pd.Series([""] * len(df), index=df.index)

# Amount = SMSs credit used × 0.14
if sms_col:
    # Coerce to numeric, fill NaN with 0, multiply
    amount = pd.to_numeric(df[sms_col], errors="coerce").fillna(0) * 0.14
else:
    amount = pd.Series([0] * len(df), index=df.index)

# --- 4) Ask for Reference ---
reference_value = input("\nEnter Reference (applies to all rows): ").strip()

# --- 5) Build final accounting DataFrame ---
final_cols = [
    "TxDate",
    "Description",
    "Reference",
    "Amount",
    "UseTax",
    "TaxType",
    "TaxAccount",
    "TaxAmount",
    "Project",
    "Account",
    "IsDebit",
]

out_df = pd.DataFrame({
    "TxDate": txdate,
    "Description": description,
    "Reference": reference_value,
    "Amount": amount.round(2),   # round to 2 decimals
    "UseTax": "N",
    "TaxType": "",
    "TaxAccount": "",
    "TaxAmount": 0,
    "Project": "",
    "Account": "8002/BLM/027/046",
    "IsDebit": "N",
})[final_cols]

# --- 6) Preview ---
print("\n=== Final columns ===")
print(list(out_df.columns))
print("\n=== Final preview (first 10 rows) ===")
print(out_df.head(10).to_string(index=False))

# --- 7) Save next to source ---
base, _ = os.path.splitext(csv_path)
out_path = base + "_accounting.csv"
out_df.to_csv(out_path, index=False)
print(f"\nSaved: {out_path}")
