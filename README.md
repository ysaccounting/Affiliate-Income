# Affiliate Income Worksheet & Balance Sheet Reconciliation

Web app: upload a QuickBooks **Consolidated P&L** export (`.xlsx`, `.xlsm`, or `.csv`)
and download the **Affiliate Income Worksheet** — one row per affiliate broker showing
year-to-date net income, Y&S's ownership share, and the new journal-entry difference.

The app does the paste-and-lookup steps for you: it reads the raw P&L, finds each
affiliate's column and the **Net Income** row, applies the fixed ownership percentages,
and writes the live worksheet formulas.

## What it does

1. **Find the affiliate columns.** It locates the entity-header row in the P&L
   (`YSM Tickets`, `YSS Tickets`, …) and the **`Net Income`** row.
2. **Net income per broker.** For each configured broker, the matching entity column's
   Net Income becomes **Total YTD Net Inc thru Last Month** (column B).
3. **Ownership & share.** Column C is the fixed **Y&S % Ownership**; column D is the live
   formula **`=B*C`** (Y&S's share); column F is **`=D-E`** (the new journal-entry amount).
4. **Date.** Column G is the month-end of the last month in the P&L period
   (`January-June, 2026` → `6/30/2026`), auto-detected and overridable in the UI.

### Output workbook

Two tabs:

- **Journal Entries** — the worksheet, one row per broker, with columns:
  `Broker` · `Total YTD Net Inc thru Last Month` · `Y&S % Ownership` ·
  `Y&S Share of Net Inc thru Last Month` · `Y&S Share of Net Inc already in the P&L` ·
  `Difference - New Journal Entry` · `Date`
- **Consolidated P&L** — the uploaded report, embedded for reference (copied with its
  formatting for `.xlsx` / `.xlsm`, or as a plain values dump for `.csv`). Rows through
  the entity-header row are frozen so the company names stay visible while scrolling.

On the Journal Entries tab, negative dollar amounts render in **red**.

**Column E** (`…already in the P&L`) is read from the **K-1 income** lines booked on the
YS Affiliates LLC entity (the `K-1 - …` rows). That K-1 income lags one month — it is YTD
through two months ago, not last month — so column F (`=D-E`) is the incremental journal
entry for the latest month. The K-1 row label for each broker is set in the `BROKERS`
table, and the value is read from the `YS Affiliates` column (falling back to `Total`).

## Balance Sheet Reconciliation

The second upload zone takes the **Consolidated Balance Sheet** QBO export and checks
that each `Inv - …` investment account on **YS Affiliates LLC's** books ties to the
affiliate's own equity. For each broker:

```
Expected Inv Balance = Total for Equity - Y&S            (on the affiliate's column)
                     + Y&S% x (Retained Earnings + Net Income)   (on the affiliate's column)

Difference = Inv Balance per Books - Expected Inv Balance
```

`Total for Equity - Y&S` is already net of capital contributions less distributions
(QBO reports distributions as negative), so it adds directly — no sign flip needed.

The output workbook has two tabs:

- **Reconciliation** — one row per broker: Inv Balance per Books, Total for
  Equity - Y&S, Retained Earnings, Net Income, Y&S % Ownership, Y&S Share of
  RE + Net Inc (`=F*(D+E)`), Expected Inv Balance (`=C+G`), Difference (`=B-H`),
  and a Status of `OK` / `Review`. A live TOTAL row sums the money columns.
- **Consolidated Balance Sheet** — the uploaded report, trimmed to just the rows the
  reconciliation reads: the report header rows, the `Inv - …` accounts, Retained
  Earnings, Net Income, and Total for Equity - Y&S.

Rows flagged `Review` are genuine variances to chase — the tool surfaces them, it does
not adjust them.

## Broker configuration

The broker list, each broker's P&L column name, and ownership % live in the `BROKERS`
table at the top of `app.py`:

```python
BROKERS = [
    ("YSM",         "YSM Tickets",         0.50, "K-1 - YSM (Grossman)", "Inv - YSM (Grossman)"),
    ("YSS",         "YSS Tickets",         0.50, "K-1 - YSS (Sternbuch)", "Inv - YSS (Sternbuch)"),
    ("YSKG",        "YSKG Tickets",        0.25, "K-1 - YSKG",           "Inv - YSKG"),
    ("YSTL",        "YS TL Tickets",       0.35, "K-1 - YS TL",          "Inv - YS TL"),
    # (display, P&L/BS entity column, ownership, K-1 row label, Inv account label)
    ...
]
```

Columns are matched to the P&L by (normalized) header name, so column order in the export
does not matter. If a broker's column is missing from a given P&L, the row is still written
(blank net income) and the app warns you. If the P&L has a new `… Tickets` column that
isn't set up as a broker, the app flags it so you can add it.

## Input format

The app expects the **raw** QuickBooks Consolidated P&L export:

- A row of **entity headers** (one column per affiliate, plus parent / roll-up / total
  columns, which are ignored).
- A **`Net Income`** row (distinct from `Net Operating Income` / `Net Other Income`).
- A **period line** near the top (e.g. `January-June, 2026`) used to auto-detect the date.

## Run locally

```bash
pip install -r requirements.txt
python app.py        # http://localhost:5000
```

## Deploy: GitHub → Railway

1. Push this folder to a GitHub repo.
2. Railway → New Project → Deploy from GitHub repo → pick it.
3. Railway auto-detects Python (Nixpacks) and uses the start command in `railway.json`.
   No env vars needed; `$PORT` is provided automatically.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask backend — P&L parsing, broker lookup, workbook builder |
| `index.html` | Single-page upload UI |
| `requirements.txt` | Python dependencies |
| `Procfile` / `railway.json` | Start command for Railway |
