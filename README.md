# Project Recon

A self-contained project exploring treasury operations automation with
Python. It covers three capabilities end to end:

| Capability | Implementation |
| --- | --- |
| Reconciliation support | `recon.py` automatically matches bank statement vs internal ledger |
| Reporting automation | one-click formatted Excel report |
| Data checks | automatic check that the total variance is fully explained |

## Project structure

```
treasury-recon-practice/
├── generate_sample_data.py   # generates sample data (with planted discrepancies)
├── recon.py                  # main script: matching + report + data check + email
├── recon_fx.py               # advanced: converts everything to HKD, then reconciles
├── requirements.txt
├── README.md
├── data/                     # sample CSVs (auto-generated)
│   ├── bank_statement.csv    # bank statement
│   ├── internal_ledger.csv   # internal ledger
│   └── fx_rates.csv          # FX rate table (HKD base)
└── output/                   # Excel reports (auto-generated)
```

## Quick start

```bash
cd treasury-recon-practice
python3 generate_sample_data.py   # generate sample data
python3 recon.py                  # run reconciliation and build the Excel report
python3 recon_fx.py               # advanced: convert everything to HKD and reconcile
```

Open the xlsx in `output/` (base report: 6 sheets — Summary, Matched, Mismatches, Bank Only, Ledger Only, Duplicates; the FX report adds an FX Rates sheet).

> Tip: in a real project, isolate dependencies with a virtual environment:
> `python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt`

## Matching logic (what recon.py does)

1. **Clean**: uppercase references, strip whitespace, round amounts to 2dp (real data always has formatting inconsistencies)
2. **Step 1**: reference + amount + currency identical → matched
3. **Step 2**: same reference but amount/currency differs → mismatches
4. **Step 3**: remaining rows matched by date+currency+amount (simulates the bank using its own reference numbering)
5. **Duplicate references** → flagged for review, never auto-matched (real systems behave the same way — duplicates must be confirmed by a human)
6. **Data check**: the variance between the two totals should equal the sum of the classified differences; if it does not, something has been missed

## What each file does

- `generate_sample_data.py`: fixed random seeds (reproducibility), dict ↔ DataFrame conversion, date handling
- `recon.py`:
  - **pandas**: `merge`, `groupby`, `value_counts`, `isin`, boolean indexing
  - **openpyxl**: multiple sheets, header colours, borders, thousand-separator number formats, frozen top row, auto-filters
  - **data checks**: total variance vs explained variance
  - **smtplib**: email with attachment (dry-run by default — never sends for real unless you ask it to)
- `recon_fx.py`:
  - converting to a base currency (HKD) with an FX rate table
  - validating that every currency has a rate before doing any math
  - tolerance matching for FX rounding
  - reusing helpers from `recon.py` via `import recon`

## The FX variant (recon_fx.py)

The first extension, implemented as a standalone script. It covers the
core scenario of a multi-currency broker: reconciling in one base currency (HKD).

What it adds on top of `recon.py`:

1. **FX rate table** — `data/fx_rates.csv` with a `rate_to_hkd` per currency; the script fails fast if a currency in the data has no rate (a real data check)
2. **Conversion** — every amount becomes `amount_hkd = amount × rate`, while the original currency and amount are kept in the report
3. **Tolerance matching** — small HKD differences (≤ 0.05) are treated as FX rounding and match with `match_type = "FX tolerance"` (real systems do this; otherwise every rounding cent shows up as a break)
4. **Stale-rate simulation** — the script deliberately nudges 2 EUR rows on the ledger side by a few cents, as if the ledger had used a slightly different rate source, so the tolerance matching has something to catch
5. **FX Rates sheet** — the report includes a sheet showing the rate applied per currency and how much each side converted to

Play with it: change `TOLERANCE_HKD`, add a new currency + rate to `fx_rates.csv`, or add your own stale-rate nudge.

## Ideas to extend

1. ~~Single-currency reconciliation~~ **implemented in `recon_fx.py`.** Try changing `TOLERANCE_HKD`, or add a new currency + rate to `fx_rates.csv`
2. ~~Tolerance matching~~ **demonstrated in `recon_fx.py`.** Try adding the same tolerance logic to `recon.py` itself (it currently matches exactly)
3. **Business days / settlement dates**: use the `holidays` library to check whether settlement dates fall on HK public holidays and whether T+2 was calculated correctly
4. **Scheduling**: `crontab -e` to run `python3 recon.py` at 9am every day, unattended
5. **More data checks**: large-transaction alerts (flag amounts above a threshold), negative balances, gaps in reference numbering
6. **Really send the email**: fill in `EMAIL_CONFIG` at the top of `recon.py` and run with `--send-email`
7. **Iterate with Claude Code**: hand any of the above to the AI, let it change the script, and verify the results against the data yourself

## Project highlights

- **Multi-step matching**: reference-based matching first, then a date+amount fallback, with duplicate references held back for manual review
- **Variance data check**: proves the books balance by explaining the total variance between the two sides, and points to specific records when it doesn't
- **FX conversion** (`recon_fx.py`): FX rate table → conversion to a single base currency (HKD) → tolerance matching for rounding differences
- **AI-assisted workflow**: AI wrote the code, and every result was verified against the sample data by hand

## Reading guide

Read through `recon.py` first, then:

1. Edit `data/*.csv` by hand (delete a row, change an amount, change a reference to lowercase) and rerun `python3 recon.py` — watch how the report changes
2. Ask Claude Code about anything unclear: "explain why recon.py does X in step Y"
3. If something breaks, share the error with Claude Code and ask it to explain the cause
