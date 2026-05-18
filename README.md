# Product Comparison Tool

Streamlit app for comparing the financial outcome of two electricity pricing products on a Belgian industrial client's consumption profile.

---

## What it does

Computes the annual and monthly cost difference between:

- **Weighted Average (WA)** — each hour is priced individually: `price_h = Coeff × Belpex_h + Adder`
- **Arithmetic Average (AA)** — each month is priced at the monthly average: `price = Coeff × avg(Belpex) + Adder`

The delta between the two depends entirely on the client's load shape relative to the intra-month Belpex price curve. A client consuming more during high-price hours pays more on WA than AA; a client concentrated in off-peak hours pays less.

Both products support independent peak / off-peak coefficients and adders.

---

## Input file format

3-column CSV or Excel file:

| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Date | Time | Value (kW or kWh) |

- Supported date formats: `YYYY-MM-DD`, `DD.MM.YYYY`, `DD/MM/YYYY`
- Supported granularity: 15-min QH or hourly (auto-detected)
- Supported encodings: UTF-8, Latin-1
- Supported separators: comma, semicolon, tab

---

## Product parameters

| Field | Label in UI | Meaning |
|-------|-------------|---------|
| `a_p` | Index coeff. (peak) | Multiplier on Belpex for peak hours. Typical: 0.90–1.10 |
| `b_p` | Fixed adder (peak) | Fixed €/MWh add-on for peak hours. Typical: −5 to +5 |
| `a_d` | Index coeff. (off-peak) | Multiplier on Belpex for off-peak hours |
| `b_d` | Fixed adder (off-peak) | Fixed €/MWh add-on for off-peak hours |

**Peak hours:** Monday–Friday 08:00–19:59 Brussels local time.  
**Off-peak hours:** all other hours (nights, weekends, public holidays not excluded).

---

## Baseload Click Sensitivity

Simulates locking a fixed volume of the WA product at a fixed price (the "clicked" price), creating a mixed product:

- **Baseload slice** — fixed MWh/month (anchored to the minimum monthly volume × selected %), priced at the clicked price. Fully decoupled from Belpex.
- **Swing slice** — remaining volume, stays on the full floating WA rate.

Outputs:
- Monthly Cost Decomposition — stacked bars showing baseload (dark navy) vs swing (blue) vs AA reference
- Cost Convergence Curve — how total mixed cost varies from 0% to 100% baseload click
- Spread Sensitivity — three scenarios (Bear ×0.5, Historical ×1.0, Bull ×2.0) using the formula:

```
Belpex_adj(h) = μ_month + k × (Belpex(h) − μ_month)
```

Monthly averages are preserved in all scenarios; only within-month hourly variance changes. The AA product cost is therefore identical across all three curves by construction.

---

## Belpex data

Static CSV sourced from EPEX Spot Belgium day-ahead prices (from 2015).  
Path: `../03. Energy Analysis Platform/data/Day ahead Belgium from 2015.csv`

The app shows the last available date on load. If the consumption file extends beyond the last Belpex date, missing hours are filled with the period median (flagged with a warning).

---

## How to run

```powershell
& "c:\Users\DEBXA\OneDrive - Luminus\04. Knowledge Database\Python\venv\Scripts\python.exe" -m streamlit run app.py
```

Requires the shared `venv` in the parent Python folder. Do not use the bare `streamlit` CLI — it is blocked by company policy.

---

## Output

The **Download Excel Report** button exports a `.xlsx` file named `{source_file}_{date}_comparison.xlsx` with three sheets:

| Sheet | Contents |
|-------|----------|
| Summary | Total cost per product, delta, volume |
| Monthly Breakdown | Month-by-month volumes, Belpex averages, effective prices, costs, delta |
| Parameters | Product types, peak/off-peak formulas, export date — full audit trail |

---

## Known limitations

- Belpex dataset is static (not live). Refresh the CSV periodically.
- Peak definition does not exclude Belgian public holidays.
- Spread Sensitivity is a linear mean-preserving rescaling, not a stochastic simulation. It does not model load-price correlation changes across scenarios.
- No sidebar; all controls are in the main area for client presentation use.
