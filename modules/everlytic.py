# modules/everlytic.py
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd


class Everlytic:
    EXPORT_COLS = [
        "TxDate","Description","Reference","Amount","UseTax",
        "TaxType","TaxAccount","TaxAmount","Project","Account","IsDebit",
    ]
    ACCOUNT_DEFAULT = "8002/BLM/027/046"

    def __init__(self, parent):
        self.container = ttk.Frame(parent)
        self.container.pack(fill="both", expand=True)
        self.df_preview = None
        self.source_path = None

        ttk.Label(self.container, text="📑 Everlytic CSV/Excel Report → Accounting", font=("Arial", 16, "bold")).pack(pady=(8, 2))
        ttk.Label(self.container, text="1) Select CSV/Excel  •  2) Preview  •  3) Export CSV").pack(pady=(0, 8))

        controls = ttk.Frame(self.container); controls.pack(pady=6)
        ttk.Button(controls, text="📁 Select CSV/Excel", command=self.load_file).grid(row=0, column=0, padx=8)
        ttk.Button(controls, text="⬇️ Export (CSV)", command=self.export_accounting_csv).grid(row=0, column=1, padx=8)

        table_wrap = ttk.Frame(self.container); table_wrap.pack(fill="both", expand=True, pady=8)
        self.tree = ttk.Treeview(table_wrap, show="headings"); self.tree.pack(side="left", fill="both", expand=True)
        yscroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview); yscroll.pack(side="right", fill="y")
        xscroll = ttk.Scrollbar(self.container, orient="horizontal", command=self.tree.xview); xscroll.pack(fill="x")
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.status = ttk.Label(self.container, text="No file loaded.", anchor="w")
        self.status.pack(fill="x", pady=(4, 2))

    # ---------- File IO ----------
    def load_file(self):
        path = filedialog.askopenfilename(
            title="Select Everlytic file",
            filetypes=[("CSV/Excel files", "*.csv *.xlsx *.xls"), ("CSV", "*.csv"), ("Excel", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            # Detect by extension
            ext = os.path.splitext(path)[1].lower()
            if ext in [".xlsx", ".xls"]:
                df = pd.read_excel(path, header=1, engine="openpyxl")
            else:
                df = pd.read_csv(path, header=1)

            df.columns = [str(c).strip() for c in df.columns]

            # 1) Always drop the first data row after the headers (banner line)
            if len(df) > 0:
                df = df.iloc[1:].reset_index(drop=True)

            # 2) ALSO remove any stray branch banners
            df, removed_banners = self._drop_banner_rows(df)

            # 3) Remove any reprinted header lines
            df, removed_repeat = self._drop_repeated_header_rows(df)

            # Build preview...
            n = len(df)
            empty_series = pd.Series([""] * n, index=df.index)
            zero_series = pd.Series([0] * n, index=df.index)

            description = (
                "Everlytic - " + df["Message Subject"].astype(str)
                if "Message Subject" in df.columns else empty_series
            )
            txdate = (
                pd.to_datetime(df["Send Date"], errors="coerce").dt.strftime("%d/%m/%Y")
                if "Send Date" in df.columns else empty_series
            )
            sms_credits = (
                pd.to_numeric(df["SMSs credit used"], errors="coerce").fillna(0)
                if "SMSs credit used" in df.columns else zero_series
            )
            amount = (sms_credits * 0.14).round(2)

            sms_sent = (
                df["Sms Sent"] if "Sms Sent" in df.columns else
                (df["SMSs sent"] if "SMSs sent" in df.columns else empty_series)
            )

            self.df_preview = pd.DataFrame({
                "TxDate": txdate,
                "Description": description,
                "SMSs credit used": sms_credits,
                "Sms Sent": sms_sent,
                "Amount": amount,
            })

            self._display(self.df_preview)
            self.source_path = path
            self.status.config(
                text=(
                    f"Loaded: {os.path.basename(path)} | {len(self.df_preview)} rows "
                    f"(removed {removed_banners} banner line(s), {removed_repeat} repeated header line(s))"
                )
            )

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{e}")



    # ---------- Export ----------
    def export_accounting_csv(self):
        if self.df_preview is None or self.df_preview.empty:
            messagebox.showwarning("No Data", "Load a CSV first.")
            return

        ref = simpledialog.askstring("Reference", "Enter Reference (applies to all rows):")
        if ref is None:
            return

        path = filedialog.asksaveasfilename(
            title="Save Accounting CSV",
            defaultextension=".csv",
            filetypes=[("CSV (UTF-8)", "*.csv")],
        )
        if not path:
            return

        try:
            out_df = self._build_export_dataframe(reference_value=ref)
            out_df.to_csv(path, index=False)
            messagebox.showinfo("Exported", f"Saved to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export:\n{e}")

    # ---------- Transform ----------
    def _build_export_dataframe(self, reference_value: str) -> pd.DataFrame:
        src = self.df_preview.copy()

        base = pd.DataFrame({
            "TxDate": src["TxDate"],
            "Description": src["Description"],
            "Reference": reference_value,
            "Amount": src["Amount"],
            "UseTax": "N",
            "TaxType": "",
            "TaxAccount": "",
            "TaxAmount": 0,
            "Project": "",
            "Account": self.ACCOUNT_DEFAULT,  # 8002/BLM/027/046
            "IsDebit": "N",
        })

        dup = base.copy()
        dup["IsDebit"] = "Y"
        dup["Account"] = "3020/BLM/"

        out = pd.concat([base, dup], ignore_index=True)
        return out[self.EXPORT_COLS]

    # ---------- Helpers ----------
    
    


    # ---------- Display ----------
    def _display(self, df: pd.DataFrame):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(df.columns)
        for col in df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=max(120, min(320, len(col) * 12)), anchor="w")
        for _, row in df.fillna("").iterrows():
            self.tree.insert("", "end", values=list(row))
