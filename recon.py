"""
Mini reconciliation script (core of the practice project)

Flow:
  1. Read the bank statement and internal ledger CSVs; clean and normalise fields
  2. Match in three steps:
     a. reference + amount + currency identical          -> matched
     b. same reference but amount (or currency) differs  -> mismatch
     c. remaining rows matched on date+currency+amount   -> matched (no shared reference)
  3. Unmatched rows go to "bank only" and "ledger only"
  4. References appearing more than once on either side go to "duplicates" for review
  5. Data check: is the variance between the two totals fully explained?
  6. Write a formatted Excel report + console summary
  7. Optional: email the report via SMTP (dry-run by default)

Usage:
    python3 recon.py
    python3 recon.py --send-email     # only sends once EMAIL_CONFIG is filled in
"""

import argparse
import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------
# Configuration: fill in your SMTP settings here to practise emailing.
# Never hardcode the password in the code — use the
# RECON_EMAIL_PASSWORD environment variable instead.
# (Gmail requires an "app password", not your login password.)
# ---------------------------------------------------------------
EMAIL_CONFIG = {
    "smtp_server": "",                 # e.g. "smtp.gmail.com"
    "smtp_port": 587,
    "username": "",                    # sender email address
    "password": "",                    # leave empty to read RECON_EMAIL_PASSWORD env var
    "from_addr": "",
    "to_addrs": [],                    # e.g. ["team@example.com"]
}

BASE = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
OUTPUT_DIR = BASE / "output"


# ================= 1. Core reconciliation logic =================

def reconcile(df_bank, df_ledger):
    """Three-step matching; returns the classified DataFrames and summary stats."""

    # ---- Clean: uppercase references, strip whitespace, round amounts to 2dp ----
    # (Real data always has formatting inconsistencies — this step is essential.)
    for df in (df_bank, df_ledger):
        df["reference"] = df["reference"].fillna("").str.strip().str.upper()
        df["amount"] = df["amount"].round(2)

    # ---- Count how many times each reference appears on each side ----
    b_cnt = df_bank["reference"].value_counts()
    l_cnt = df_ledger["reference"].value_counts()
    shared = set(b_cnt.index) & set(l_cnt.index)

    # References appearing exactly once on both sides -> safe to compare automatically
    clean_refs = {r for r in shared if b_cnt[r] == 1 and l_cnt[r] == 1}
    # References appearing more than once on either side -> flag for manual review
    # (real systems do not auto-match duplicates either)
    dup_refs = {r for r in shared if b_cnt[r] > 1 or l_cnt[r] > 1}

    # ---- Step 1: unique references on both sides -> compare amount and currency ----
    b1 = df_bank[df_bank["reference"].isin(clean_refs)]
    l1 = df_ledger[df_ledger["reference"].isin(clean_refs)]
    m1 = b1.merge(l1, on="reference", suffixes=("_bank", "_ledger"))

    cur_ok = m1["currency_bank"] == m1["currency_ledger"]
    amt_ok = m1["amount_bank"] == m1["amount_ledger"]
    ok = cur_ok & amt_ok

    matched = pd.DataFrame({
        "date":            m1.loc[ok, "date_bank"],
        "reference_bank":  m1.loc[ok, "reference"],
        "reference_ledger": m1.loc[ok, "reference"],
        "description":     m1.loc[ok, "description_bank"],
        "counterparty":    m1.loc[ok, "counterparty_bank"],
        "currency":        m1.loc[ok, "currency_bank"],
        "amount":          m1.loc[ok, "amount_bank"],
        "match_type":      "reference+amount",
    })

    mismatch = pd.DataFrame({
        "date":            m1.loc[~ok, "date_bank"],
        "reference":       m1.loc[~ok, "reference"],
        "description":     m1.loc[~ok, "description_bank"],
        "counterparty":    m1.loc[~ok, "counterparty_bank"],
        "currency_bank":   m1.loc[~ok, "currency_bank"],
        "currency_ledger": m1.loc[~ok, "currency_ledger"],
        "amount_bank":     m1.loc[~ok, "amount_bank"],
        "amount_ledger":   m1.loc[~ok, "amount_ledger"],
    })
    mismatch["difference"] = np.where(
        mismatch["currency_bank"] == mismatch["currency_ledger"],
        (mismatch["amount_bank"] - mismatch["amount_ledger"]).round(2),
        np.nan,  # different currencies cannot be netted — leave blank for manual review
    )

    # ---- Duplicate references -> flagged for review ----
    dup = pd.concat([
        df_bank[df_bank["reference"].isin(dup_refs)].assign(side="bank"),
        df_ledger[df_ledger["reference"].isin(dup_refs)].assign(side="ledger"),
    ], ignore_index=True)
    dup = dup[["side", "date", "reference", "description",
               "counterparty", "currency", "amount"]]

    # ---- Step 2: remaining rows (reference on one side only) matched by
    #      date+currency+amount ----
    # Mimics real life: the bank uses its own reference numbering, which does
    # not match internal references.
    # Only auto-match date+currency+amount combos that occur exactly once on
    # each side; otherwise leave them as one-sided differences rather than
    # force-pairing coincidences.
    def _key(df):
        return (df["date"].astype(str) + "|" + df["currency"] + "|"
                + df["amount"].map(lambda x: f"{x:.2f}"))

    b_rem = df_bank[~df_bank["reference"].isin(shared)].copy()
    l_rem = df_ledger[~df_ledger["reference"].isin(shared)].copy()
    b_rem["_key"] = _key(b_rem)
    l_rem["_key"] = _key(l_rem)

    b_n = b_rem.groupby("_key").size()
    l_n = l_rem.groupby("_key").size()
    ok_keys = set(b_n[b_n == 1].index) & set(l_n[l_n == 1].index)

    b2 = b_rem[b_rem["_key"].isin(ok_keys)]
    l2 = l_rem[l_rem["_key"].isin(ok_keys)]
    m2 = b2.merge(l2, on="_key", suffixes=("_bank", "_ledger"))

    matched2 = pd.DataFrame({
        "date":            m2["date_bank"],
        "reference_bank":  m2["reference_bank"],
        "reference_ledger": m2["reference_ledger"],
        "description":     m2["description_bank"],
        "counterparty":    m2["counterparty_bank"],
        "currency":        m2["currency_bank"],
        "amount":          m2["amount_bank"],
        "match_type":      "date+amount",
    })

    # ---- Remainder -> one-sided differences ----
    cols = ["date", "reference", "description", "counterparty", "currency", "amount"]
    bank_only = b_rem[~b_rem["_key"].isin(ok_keys)][cols].copy()
    ledger_only = l_rem[~l_rem["_key"].isin(ok_keys)][cols].copy()

    matched_all = pd.concat([matched, matched2], ignore_index=True)

    # ---- Summary + data check ----
    summary = {
        "bank_count": len(df_bank),
        "ledger_count": len(df_ledger),
        "matched_count": len(matched),
        "matched2_count": len(matched2),
        "mismatch_count": len(mismatch),
        "bank_only_count": len(bank_only),
        "ledger_only_count": len(ledger_only),
        "dup_count": len(dup),
        "bank_total": df_bank["amount"].sum(),
        "ledger_total": df_ledger["amount"].sum(),
    }

    # Data check: the variance between the two totals should equal the
    # sum of the classified differences. If it does not, something is
    # still unclassified (e.g. a duplicate booking).
    variance = summary["bank_total"] - summary["ledger_total"]
    explained = (bank_only["amount"].sum()
                 - ledger_only["amount"].sum()
                 + mismatch["amount_bank"].fillna(0).sum()
                 - mismatch["amount_ledger"].fillna(0).sum())
    check = {
        "variance": round(variance, 2),
        "explained": round(explained, 2),
        "unexplained": round(variance - explained, 2),
    }

    return {
        "matched": matched_all,
        "mismatch": mismatch,
        "bank_only": bank_only,
        "ledger_only": ledger_only,
        "dup": dup,
        "summary": summary,
        "check": check,
    }


# ================= 2. Excel report =================

SHEET_COLORS = {
    "Summary": "1F4E79",
    "Matched": "2E7D32",
    "Mismatches": "B8860B",
    "Bank Only": "C0392B",
    "Ledger Only": "D35400",
    "Duplicates": "6C3483",
}

THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _to_str_dates(df):
    """Write dates as YYYY-MM-DD strings in Excel."""
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d")
    return df


def _col_width(values):
    """Estimate column width from content (non-ASCII chars count double)."""
    w = 0
    for v in values:
        s = str(v)
        w = max(w, sum(2 if ord(c) > 127 else 1 for c in s))
    return min(max(w + 2, 10), 36)


def _write_table(ws, df, color):
    """Write a DataFrame to a sheet with header styling, borders and number formats."""
    df = _to_str_dates(df)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor=color)

    for j, col in enumerate(df.columns, start=1):
        cell = ws.cell(row=1, column=j, value=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws.column_dimensions[get_column_letter(j)].width = _col_width(
            [col] + list(df[col].astype(str)))

    for i, row in enumerate(df.itertuples(index=False), start=2):
        for j, value in enumerate(row, start=1):
            cell = ws.cell(row=i, column=j, value=value)
            cell.border = BORDER
            if isinstance(value, float):
                cell.number_format = "#,##0.00"

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _write_summary(ws, r, color):
    """Summary sheet: key figures + data check result."""
    s, c = r["summary"], r["check"]
    balanced = c["unexplained"] == 0
    rows = [
        ("Report generated", datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("", ""),
        ("Bank statement rows", s["bank_count"]),
        ("Internal ledger rows", s["ledger_count"]),
        ("", ""),
        ("Matched (reference+amount)", s["matched_count"]),
        ("Matched (date+amount)", s["matched2_count"]),
        ("Amount mismatches", s["mismatch_count"]),
        ("Bank only (missing in ledger)", s["bank_only_count"]),
        ("Ledger only (missing in bank)", s["ledger_only_count"]),
        ("Duplicate references", s["dup_count"]),
        ("", ""),
        ("Bank total", s["bank_total"]),
        ("Ledger total", s["ledger_total"]),
        ("Variance", c["variance"]),
        ("Explained", c["explained"]),
        ("Unexplained", c["unexplained"]),
        ("", ""),
        ("Data check result",
         "✓ balanced — all variance explained" if balanced
         else "✗ unexplained variance — review Duplicates sheet"),
    ]

    ws.cell(row=1, column=1, value="Reconciliation Summary").font = Font(
        bold=True, size=14, color=color)

    for i, (label, value) in enumerate(rows, start=3):
        lc = ws.cell(row=i, column=1, value=label)
        vc = ws.cell(row=i, column=2, value=value)
        lc.font = Font(bold=True)
        if isinstance(value, float):
            vc.number_format = "#,##0.00"
        if label == "Data check result":
            lc.font = Font(bold=True, color="2E7D32" if balanced else "C0392B")
            vc.font = Font(bold=True, color="2E7D32" if balanced else "C0392B")

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 38


def build_excel_report(r, path):
    wb = Workbook()
    wb.remove(wb.active)  # drop the default empty sheet

    _write_summary(wb.create_sheet("Summary"), r, SHEET_COLORS["Summary"])
    _write_table(wb.create_sheet("Matched"), r["matched"], SHEET_COLORS["Matched"])
    _write_table(wb.create_sheet("Mismatches"), r["mismatch"], SHEET_COLORS["Mismatches"])
    _write_table(wb.create_sheet("Bank Only"), r["bank_only"], SHEET_COLORS["Bank Only"])
    _write_table(wb.create_sheet("Ledger Only"), r["ledger_only"], SHEET_COLORS["Ledger Only"])
    _write_table(wb.create_sheet("Duplicates"), r["dup"], SHEET_COLORS["Duplicates"])

    wb.save(path)


# ================= 3. Console summary =================

def format_summary(r):
    s, c = r["summary"], r["check"]
    balanced = c["unexplained"] == 0
    # computed first — f-string expressions must stay on one line in Python < 3.12
    result_msg = ("✓ balanced — all variance explained" if balanced
                  else "✗ unexplained variance — review Duplicates sheet")
    lines = [
        "=" * 60,
        f"RECONCILIATION REPORT  {datetime.now():%Y-%m-%d %H:%M}",
        "=" * 60,
        f"Bank statement: {s['bank_count']:>4} rows | Internal ledger: {s['ledger_count']:>4} rows",
        "",
        f"  ✓ Matched (reference+amount)     {s['matched_count']:>4}",
        f"  ✓ Matched (date+amount)          {s['matched2_count']:>4}",
        f"  ✗ Amount mismatches              {s['mismatch_count']:>4}",
        f"  ✗ Bank only (missing in ledger)  {s['bank_only_count']:>4}",
        f"  ✗ Ledger only (missing in bank)  {s['ledger_only_count']:>4}",
        f"  ⚠ Duplicate references (review)  {s['dup_count']:>4}",
        "=" * 60,
        "Data check:",
        f"  Bank total     {s['bank_total']:>14,.2f}",
        f"  Ledger total   {s['ledger_total']:>14,.2f}",
        f"  Variance       {c['variance']:>14,.2f}",
        f"  Explained      {c['explained']:>14,.2f}",
        f"  Unexplained    {c['unexplained']:>14,.2f}",
        f"  Result: {result_msg}",
    ]
    return "\n".join(lines)


# ================= 4. Email the report =================

def send_email(report_path, summary_text, dry_run=True):
    cfg = EMAIL_CONFIG
    if not cfg["smtp_server"]:
        print("\n[Email] SMTP not configured — skipping. "
              "(Fill in EMAIL_CONFIG at the top of recon.py to enable.)")
        return

    password = cfg["password"] or os.environ.get("RECON_EMAIL_PASSWORD", "")
    if not password:
        print("\n[Email] No password set (EMAIL_CONFIG or the RECON_EMAIL_PASSWORD env var).")
        return

    msg = EmailMessage()
    msg["Subject"] = f"Reconciliation report {datetime.now():%Y-%m-%d}"
    msg["From"] = cfg["from_addr"] or cfg["username"]
    msg["To"] = ", ".join(cfg["to_addrs"])
    msg.set_content(f"Please find today's reconciliation report attached.\n\n{summary_text}")

    with open(report_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=report_path.name,
        )

    if dry_run:
        print(f"\n[Email] DRY-RUN: not actually sent (pass --send-email to send for real)")
        print(f"  To:      {msg['To']}")
        print(f"  Subject: {msg['Subject']}")
        return

    with smtplib.SMTP(cfg["smtp_server"], cfg["smtp_port"]) as s:
        s.starttls()
        s.login(cfg["username"], password)
        s.send_message(msg)
    print("\n[Email] Sent")


# ================= Main =================

def main():
    parser = argparse.ArgumentParser(description="Mini reconciliation script (practice project)")
    parser.add_argument("--data-dir", default=str(DATA_DIR), help="CSV data directory")
    parser.add_argument("--output-dir", default=str(OUTPUT_DIR), help="report output directory")
    parser.add_argument("--send-email", action="store_true",
                        help="send the email for real (requires EMAIL_CONFIG)")
    args = parser.parse_args()

    data_dir = Path(args.data_dir)
    bank_path = data_dir / "bank_statement.csv"
    ledger_path = data_dir / "internal_ledger.csv"
    for p in (bank_path, ledger_path):
        if not p.exists():
            raise SystemExit(f"Missing {p} — run `python3 generate_sample_data.py` first")

    print("Reading data...")
    df_bank = pd.read_csv(bank_path, parse_dates=["date"])
    df_ledger = pd.read_csv(ledger_path, parse_dates=["date"])

    r = reconcile(df_bank, df_ledger)

    out_dir = Path(args.output_dir)
    out_dir.mkdir(exist_ok=True)
    report_path = out_dir / f"recon_report_{datetime.now():%Y%m%d_%H%M}.xlsx"
    build_excel_report(r, report_path)

    print(format_summary(r))
    print(f"\nReport saved: {report_path}")
    send_email(report_path, format_summary(r), dry_run=not args.send_email)


if __name__ == "__main__":
    main()
