"""
FX conversion reconciliation (advanced exercise #1)

Converts every transaction on both sides into HKD using an FX rate
table, then reconciles in HKD terms — the core scenario of a
multi-currency broker.

What this adds on top of recon.py:
  1. Loads data/fx_rates.csv and validates that every currency in the
     data has a rate (a real data check — fail fast, before any math)
  2. Converts both sides to HKD, keeping the original currency and
     amount columns in the report
  3. Matches with an FX tolerance: small HKD differences (<= 0.05) are
     treated as FX rounding rather than real breaks
  4. Simulates a stale rate on the ledger side for 2 EUR rows so the
     tolerance matching has something to catch
  5. Adds an "FX Rates" sheet showing the rate applied per currency and
     how much each side converted to

Usage:
    python3 recon_fx.py
    python3 recon_fx.py --tolerance 0.10
    python3 recon_fx.py --send-email
"""

import argparse
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook

import recon  # reuse the report/email helpers from the base exercise

# Small HKD differences at or below this are treated as FX rounding.
TOLERANCE_HKD = 0.05

# References that already carry a planted discrepancy in the sample data
# (3 amount mismatches + 2 duplicate bookings) — leave them alone so this
# exercise stays predictable.
EXCLUDED_REFS = {
    "TX20261006", "TX20261016", "TX20261026",  # the 3 amount mismatches
    "TX20261011", "TX20261031",                # the 2 duplicate bookings
}


def convert(df, rates):
    """Add amount_orig and amount_hkd columns to a side's DataFrame.

    The original currency and amount are kept so the report still shows
    what the raw data looked like.
    """
    df["amount_orig"] = df["amount"]
    df["amount_hkd"] = (df["amount"] * df["currency"].map(rates)).round(2)


def simulate_stale_rate(df_ledger):
    """Nudge 2 EUR rows on the ledger side by a few cents, as if the
    ledger had used a slightly different rate source for those rows.
    This gives the tolerance matching something to catch.
    """
    eur = df_ledger[(df_ledger["currency"] == "EUR")
                    & ~df_ledger["reference"].isin(EXCLUDED_REFS)]
    nudges = [0.03, -0.03]  # sums to zero, so the overall balance is untouched
    for (idx, _), nudge in zip(eur.head(len(nudges)).iterrows(), nudges):
        df_ledger.at[idx, "amount_hkd"] = round(
            df_ledger.at[idx, "amount_hkd"] + nudge, 2)
        print(f"  [Simulated] ledger used a stale FX rate on "
              f"{df_ledger.at[idx, 'reference']} ({nudge:+.2f} HKD)")
    return df_ledger


def reconcile_fx(df_bank, df_ledger, tolerance):
    """Same three-step logic as recon.py, but in HKD and with a tolerance."""

    for df in (df_bank, df_ledger):
        df["reference"] = df["reference"].fillna("").str.strip().str.upper()
        df["amount_hkd"] = df["amount_hkd"].round(2)

    # ---- Count how many times each reference appears on each side ----
    b_cnt = df_bank["reference"].value_counts()
    l_cnt = df_ledger["reference"].value_counts()
    shared = set(b_cnt.index) & set(l_cnt.index)

    # References appearing exactly once on both sides -> safe to compare
    clean_refs = {r for r in shared if b_cnt[r] == 1 and l_cnt[r] == 1}
    # References appearing more than once on either side -> manual review
    dup_refs = {r for r in shared if b_cnt[r] > 1 or l_cnt[r] > 1}

    # ---- Step 1: unique references on both sides -> compare HKD amounts ----
    b1 = df_bank[df_bank["reference"].isin(clean_refs)]
    l1 = df_ledger[df_ledger["reference"].isin(clean_refs)]
    m1 = b1.merge(l1, on="reference", suffixes=("_bank", "_ledger"))

    cur_ok = m1["currency_bank"] == m1["currency_ledger"]
    diff = (m1["amount_hkd_bank"] - m1["amount_hkd_ledger"]).abs()
    ok = cur_ok & (diff <= tolerance)

    matched = pd.DataFrame({
        "date":            m1.loc[ok, "date_bank"],
        "reference_bank":  m1.loc[ok, "reference"],
        "reference_ledger": m1.loc[ok, "reference"],
        "description":     m1.loc[ok, "description_bank"],
        "counterparty":    m1.loc[ok, "counterparty_bank"],
        "currency":        m1.loc[ok, "currency_bank"],
        "amount_orig":     m1.loc[ok, "amount_bank"],
        "amount_hkd":      m1.loc[ok, "amount_hkd_bank"],
        # within tolerance but not exactly equal -> FX rounding, not a break
        "match_type":      np.where(diff[ok] == 0, "reference+amount", "FX tolerance"),
    })

    mismatch = pd.DataFrame({
        "date":              m1.loc[~ok, "date_bank"],
        "reference":         m1.loc[~ok, "reference"],
        "description":       m1.loc[~ok, "description_bank"],
        "counterparty":      m1.loc[~ok, "counterparty_bank"],
        "currency_bank":     m1.loc[~ok, "currency_bank"],
        "currency_ledger":   m1.loc[~ok, "currency_ledger"],
        "amount_orig_bank":  m1.loc[~ok, "amount_bank"],
        "amount_orig_ledger": m1.loc[~ok, "amount_ledger"],
        "amount_hkd_bank":   m1.loc[~ok, "amount_hkd_bank"],
        "amount_hkd_ledger": m1.loc[~ok, "amount_hkd_ledger"],
    })
    mismatch["difference_hkd"] = np.where(
        mismatch["currency_bank"] == mismatch["currency_ledger"],
        (mismatch["amount_hkd_bank"] - mismatch["amount_hkd_ledger"]).round(2),
        np.nan,  # different currencies cannot be netted — leave blank for manual review
    )

    # ---- Duplicate references -> flagged for review ----
    dup = pd.concat([
        df_bank[df_bank["reference"].isin(dup_refs)].assign(side="bank"),
        df_ledger[df_ledger["reference"].isin(dup_refs)].assign(side="ledger"),
    ], ignore_index=True)
    dup = dup[["side", "date", "reference", "description", "counterparty",
               "currency", "amount_orig", "amount_hkd"]]

    # ---- Step 2: remaining rows matched by date+amount_hkd ----
    # (all amounts are in HKD now, so currency is no longer part of the key)
    def _key(df):
        return df["date"].astype(str) + "|" + df["amount_hkd"].map(lambda x: f"{x:.2f}")

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
        "amount_orig":     m2["amount_bank"],
        "amount_hkd":      m2["amount_hkd_bank"],
        "match_type":      "date+amount",
    })

    cols = ["date", "reference", "description", "counterparty", "currency",
            "amount_orig", "amount_hkd"]
    bank_only = b_rem[~b_rem["_key"].isin(ok_keys)][cols].copy()
    ledger_only = l_rem[~l_rem["_key"].isin(ok_keys)][cols].copy()

    matched_all = pd.concat([matched, matched2], ignore_index=True)

    # ---- Summary + data check (all amounts in HKD) ----
    summary = {
        "bank_count": len(df_bank),
        "ledger_count": len(df_ledger),
        "matched_count": len(matched),
        "matched2_count": len(matched2),
        "mismatch_count": len(mismatch),
        "bank_only_count": len(bank_only),
        "ledger_only_count": len(ledger_only),
        "dup_count": len(dup),
        "bank_total": df_bank["amount_hkd"].sum(),
        "ledger_total": df_ledger["amount_hkd"].sum(),
    }

    # Tolerance-matched rows are not exactly equal, so their small net
    # difference is part of the explanation too.
    tol_mask = ok & (diff > 0)
    tol_bank = m1.loc[tol_mask, "amount_hkd_bank"].sum()
    tol_ledger = m1.loc[tol_mask, "amount_hkd_ledger"].sum()

    variance = summary["bank_total"] - summary["ledger_total"]
    explained = (bank_only["amount_hkd"].sum()
                 - ledger_only["amount_hkd"].sum()
                 + mismatch["amount_hkd_bank"].fillna(0).sum()
                 - mismatch["amount_hkd_ledger"].fillna(0).sum()
                 + tol_bank - tol_ledger)
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


# ================= Report =================

def build_rates_sheet(df_bank, df_ledger, rates_df):
    """Per-currency summary: rate applied and how much each side converted to."""
    rows = []
    for _, rrow in rates_df.iterrows():
        cur = rrow["currency"]
        b = df_bank[df_bank["currency"] == cur]
        l = df_ledger[df_ledger["currency"] == cur]
        rows.append({
            "currency": cur,
            "rate_to_hkd": rrow["rate_to_hkd"],
            "as_of": rrow["as_of"],
            "bank_rows": len(b),
            "ledger_rows": len(l),
            "bank_total_hkd": b["amount_hkd"].sum(),
            "ledger_total_hkd": l["amount_hkd"].sum(),
        })
    return pd.DataFrame(rows)


def print_conversion(rates_sheet):
    print("FX conversion (HKD base):")
    for _, row in rates_sheet.iterrows():
        print(f"  {row['currency']:<4} rate {row['rate_to_hkd']:.4f} "
              f"(as of {row['as_of']}) | "
              f"bank {int(row['bank_rows']):>2} rows {row['bank_total_hkd']:>14,.2f} HKD | "
              f"ledger {int(row['ledger_rows']):>2} rows {row['ledger_total_hkd']:>14,.2f} HKD")


def build_excel_report_fx(r, rates_sheet, path):
    """Same 6 sheets as recon.py, plus an 'FX Rates' sheet."""
    # Reusing the helpers from recon.py — the leading underscore is only a
    # naming convention for "module-internal", it still imports fine.
    wb = Workbook()
    wb.remove(wb.active)
    c = recon.SHEET_COLORS

    recon._write_summary(wb.create_sheet("Summary"), r, c["Summary"])
    recon._write_table(wb.create_sheet("Matched"), r["matched"], c["Matched"])
    recon._write_table(wb.create_sheet("Mismatches"), r["mismatch"], c["Mismatches"])
    recon._write_table(wb.create_sheet("Bank Only"), r["bank_only"], c["Bank Only"])
    recon._write_table(wb.create_sheet("Ledger Only"), r["ledger_only"], c["Ledger Only"])
    recon._write_table(wb.create_sheet("Duplicates"), r["dup"], c["Duplicates"])
    recon._write_table(wb.create_sheet("FX Rates"), rates_sheet, "1F618D")

    wb.save(path)


# ================= Main =================

def main():
    parser = argparse.ArgumentParser(
        description="FX conversion reconciliation (advanced exercise)")
    parser.add_argument("--data-dir", default=str(recon.DATA_DIR),
                        help="CSV data directory")
    parser.add_argument("--output-dir", default=str(recon.OUTPUT_DIR),
                        help="report output directory")
    parser.add_argument("--tolerance", type=float, default=TOLERANCE_HKD,
                        help="FX tolerance in HKD")
    parser.add_argument("--send-email", action="store_true",
                        help="send the email for real (requires recon.EMAIL_CONFIG)")
    args = parser.parse_args()

    data_dir = Path(args.data_dir)
    for name in ("bank_statement.csv", "internal_ledger.csv", "fx_rates.csv"):
        p = data_dir / name
        if not p.exists():
            raise SystemExit(f"Missing {p} — run `python3 generate_sample_data.py` first")

    print("Reading data and FX rates...")
    rates_df = pd.read_csv(data_dir / "fx_rates.csv")
    rates = rates_df.set_index("currency")["rate_to_hkd"]

    df_bank = pd.read_csv(data_dir / "bank_statement.csv", parse_dates=["date"])
    df_ledger = pd.read_csv(data_dir / "internal_ledger.csv", parse_dates=["date"])

    # ---- Data checks BEFORE any conversion ----
    needed = set(df_bank["currency"]) | set(df_ledger["currency"])
    missing = needed - set(rates.index)
    if missing:
        raise SystemExit(f"Missing FX rates for: {sorted(missing)} — "
                         f"update {data_dir / 'fx_rates.csv'}")
    if (rates <= 0).any():
        raise SystemExit("fx_rates.csv contains non-positive rates — fix the file")

    convert(df_bank, rates)
    convert(df_ledger, rates)

    rates_sheet = build_rates_sheet(df_bank, df_ledger, rates_df)
    print_conversion(rates_sheet)

    print("\nSimulating a stale FX rate on the ledger side:")
    df_ledger = simulate_stale_rate(df_ledger)

    r = reconcile_fx(df_bank, df_ledger, args.tolerance)

    out_dir = Path(args.output_dir)
    out_dir.mkdir(exist_ok=True)
    report_path = out_dir / f"recon_fx_report_{datetime.now():%Y%m%d_%H%M}.xlsx"
    build_excel_report_fx(r, rates_sheet, report_path)

    print()
    print(recon.format_summary(r))

    n_tol = int((r["matched"]["match_type"] == "FX tolerance").sum())
    if n_tol:
        print(f"\nNote: {n_tol} rows matched within FX tolerance "
              f"(±{args.tolerance} HKD) — see the Matched sheet.")

    print(f"\nReport saved: {report_path}")
    recon.send_email(report_path, recon.format_summary(r), dry_run=not args.send_email)


if __name__ == "__main__":
    main()
