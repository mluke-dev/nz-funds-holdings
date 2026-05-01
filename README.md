# NZ Australasian Funds — Holdings Extractor

For a hand-picked list of New Zealand Australasian-equity funds, downloads
each fund's published portfolio holdings, extracts the equity rows, normalises
security names to canonical tickers, and produces a peer-comparison matrix
showing each fund's positioning vs. the NZX50 benchmark.

Runs entirely on GitHub Actions — no Python needed locally.

## Outputs

After each run, these files are committed back to the repo under `data/`:

| File | Purpose |
|---|---|
| `holdings_clean.csv`    | Long-format clean data — one row per fund × holding. Use for filtering/pivoting. |
| `holdings_matrix.xlsx`  | Wide-format peer matrix — tickers × funds, with NZX50 benchmark and active columns. **The main analysis file.** |
| `dashboard.html`        | Interactive sortable HTML dashboard. Open in any browser. |
| `unmatched.csv`         | Rows that couldn't be classified or matched to a ticker. **Review after each run.** |
| `extract_log.txt`       | Per-fund processing log. |
| `holdings_raw/FND*.xlsx`| Original downloaded files — your audit trail. |

## How to run it

1. **Edit `selected_funds.csv`** — add or remove rows for the funds you want to analyse. Required columns: `fund_id`, `fund_name`, `provider`, `portfolio_xlsx_url`.
2. Go to the **Actions** tab of the repo.
3. Click **Extract fund holdings** in the left sidebar.
4. Click the **Run workflow** dropdown → **Run workflow**.
5. Wait ~5 minutes. The data files appear in `data/` once the run finishes.

## How it works

```
selected_funds.csv  ────┐
                        ├──►  download each .xlsx/.csv  ──►  parse  ──►  classify rows
securities.csv      ────┘                                              (equity / cash / fx / fund / bond)
                                                                              │
                                                                              ▼
                                                                      ISIN lookup
                                                                      → name lookup
                                                                              │
                                                                              ▼
                                                                       master ticker
                                                                              │
                                              ┌───────────────────────────────┴───────┐
                                              ▼                                       ▼
                                  Australasian (NZ/AU) only         Anything unmatched → unmatched.csv
                                              │
                                              ▼
                                  data/holdings_clean.csv
                                  data/holdings_matrix.xlsx (with NZX50 benchmark)
                                  data/dashboard.html
```

### What's in `securities.csv`

The master security list. Each row is one stock with:
- `ticker` — your canonical handle (e.g. `FPH`, `MFT`, `XRO`)
- `canonical_name` — display name
- `country` — `NZ` or `AU` (drives the Australasian filter)
- `sector` — for the "By Sector" sheet and dashboard grouping
- `isin` — primary identifier; ISIN-matched rows are bulletproof
- `aliases` — pipe-separated alternative names (optional)

It's seeded with the NZX50 (from the SuperLife S&P/NZX 50 Fund) plus common ASX top-50 names and a few NZX small-caps.

### When you see unmatched rows

`data/unmatched.csv` lists every row the script couldn't process. Each entry has a `reason`:

| Reason | What to do |
|---|---|
| `classified as cash/fx/settlement/bond` | Working as intended — these aren't equity. Ignore. |
| `classified as fund` | Look-through scenario. The fund holds units in another fund. v1 doesn't recurse into these. |
| `no master match` | An equity we don't have in `securities.csv` yet. Add a row to that file (or add an alias to an existing row), commit, re-run. |

Add aliases by appending to the existing row's `aliases` column, separated by `|`. Example:
```
FPH,Fisher & Paykel Healthcare,NZ,Healthcare,NZFAPE0001S2,F&P Healthcare|Fisher Paykel
```

## The matrix sheet at a glance

```
Ticker │ Name                   │ Country │ Sector       │ NZX50 │ FND1208 │ FND1240 │ FND1208_active │ ...
FPH    │ Fisher & Paykel Health │ NZ      │ Healthcare   │ 15.91 │ 19.05   │ 17.20   │ +3.14          │
MFT    │ Mainfreight            │ NZ      │ Industrials  │  3.80 │  9.39   │  8.10   │ +5.59          │
AIA    │ Auckland Airport       │ NZ      │ Infrastruct. │  9.74 │  7.04   │  6.50   │ -2.70          │
```

Active = fund weight − NZX50 weight. Positive = overweight, negative = underweight.

## Notes

- The benchmark is the SuperLife S&P/NZX 50 Index Fund (FND19363), used as a proxy for the NZX50 — it tracks the index closely and is published in the same Disclose Register format.
- The workflow has a 20-minute timeout. With ~20 funds it should finish in 2–3 minutes.
- All raw files are kept under `data/holdings_raw/` so you have an audit trail of exactly what was on the Disclose Register at the time of each run.
