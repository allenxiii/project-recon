"""
Generate sample data for reconciliation practice:
- data/bank_statement.csv    bank statement
- data/internal_ledger.csv   internal ledger (as exported from a TMS)
- data/fx_rates.csv          FX rate table (HKD base, sample rates)

Both files come from the same set of transactions, but common
reconciliation discrepancies are deliberately planted so you can
practice the matching logic in recon.py. The random seed is fixed,
so every run produces identical data.

Usage:
    python3 generate_sample_data.py
    python3 generate_sample_data.py --count 100
"""

import argparse
import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

DATA_DIR = Path(__file__).resolve().parent / "data"

CURRENCIES = ["HKD", "USD", "EUR", "CNY"]

COUNTERPARTIES = [
    "ACME Trading Ltd", "Global FX Broker", "HSBC Settlement",
    "Standard Chartered", "Citibank NA", "JP Morgan Chase",
    "Deutsche Bank AG", "Bank of China HK",
    "Client Margin A", "Client Margin B",
]

DESCRIPTIONS = [
    "client funding", "margin payment", "fx settlement",
    "interbank transfer", "commission fee", "custody fee",
    "interest income", "trade settlement",
]

# Sample FX rates (HKD base). NOT real market rates — in a real setup
# these would be pulled daily from a rate source (API, Bloomberg, TMS, ...).
FX_RATES = [
    {"currency": "HKD", "rate_to_hkd": 1.0000, "as_of": "2026-10-01"},
    {"currency": "USD", "rate_to_hkd": 7.8000, "as_of": "2026-10-01"},
    {"currency": "EUR", "rate_to_hkd": 8.5500, "as_of": "2026-10-01"},
    {"currency": "CNY", "rate_to_hkd": 1.0900, "as_of": "2026-10-01"},
]


def random_amount(rng):
    """Random amount: positive = inflow, negative = outflow."""
    base = rng.choice([100, 500, 1000, 2500, 5000, 10000, 50000, 100000])
    cents = rng.choice([0.00, 0.15, 0.50, 0.99, 123.45, 678.90])
    sign = rng.choice([1, 1, 1, -1])  # inflows slightly more frequent, like a real account
    return round(sign * (base + cents), 2)


def make_row(rng, ref, day, currency):
    return {
        "date": day.isoformat(),
        "reference": ref,
        "description": rng.choice(DESCRIPTIONS),
        "counterparty": rng.choice(COUNTERPARTIES),
        "currency": currency,
        "amount": random_amount(rng),
    }


def main(count=60):
    rng = random.Random(42)  # fixed seed -> reproducible, easier to debug the recon script
    start = date(2026, 10, 1)
    refs = [f"TX{20261000 + i}" for i in range(1, count + 1)]

    # 1. The "real world" transactions: present on both sides
    truth = [
        make_row(rng, ref, start + timedelta(days=rng.randint(0, 40)),
                 rng.choice(CURRENCIES))
        for ref in refs
    ]

    bank_rows = [dict(r) for r in truth]
    ledger_rows = [dict(r) for r in truth]

    # 2. Plant discrepancies (the issue types you meet in real reconciliation)
    # 2a. 5 rows bank-only (missed booking in the ledger)
    for i in range(5):
        day = start + timedelta(days=rng.randint(0, 40))
        bank_rows.append(make_row(rng, f"BANKONLY{100 + i}", day,
                                  rng.choice(CURRENCIES)))

    # 2b. 5 rows ledger-only (not yet settled at the bank — timing difference)
    for i in range(5):
        day = start + timedelta(days=rng.randint(0, 40))
        ledger_rows.append(make_row(rng, f"LEDGERONLY{100 + i}", day,
                                    rng.choice(CURRENCIES)))

    # 2c. 3 rows with the same reference on both sides but a different amount (keying error)
    for ref in [refs[5], refs[15], refs[25]]:
        for row in ledger_rows:
            if row["reference"] == ref:
                row["amount"] = round(row["amount"] + rng.choice([-100, 10, 9.99]), 2)

    # 2d. 2 rows booked twice in the ledger
    for ref in [refs[10], refs[30]]:
        dup = next(r for r in ledger_rows if r["reference"] == ref)
        ledger_rows.append(dict(dup))

    # 3. Shuffle the order (real statements rarely arrive in the same order)
    #    and write out
    rng.shuffle(bank_rows)
    rng.shuffle(ledger_rows)

    DATA_DIR.mkdir(exist_ok=True)
    pd.DataFrame(bank_rows).to_csv(DATA_DIR / "bank_statement.csv", index=False)
    pd.DataFrame(ledger_rows).to_csv(DATA_DIR / "internal_ledger.csv", index=False)
    pd.DataFrame(FX_RATES).to_csv(DATA_DIR / "fx_rates.csv", index=False)

    print("Sample data generated:")
    print(f"  Bank statement   : {len(bank_rows):>3} rows -> data/bank_statement.csv")
    print(f"  Internal ledger  : {len(ledger_rows):>3} rows -> data/internal_ledger.csv")
    print(f"  FX rates         : {len(FX_RATES):>3} currencies -> data/fx_rates.csv")
    print("  Planted issues   : bank-only 5 / ledger-only 5 / amount mismatch 3 / duplicate 2")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate sample reconciliation data")
    parser.add_argument("--count", type=int, default=60, help="number of base transactions")
    args = parser.parse_args()
    main(args.count)
