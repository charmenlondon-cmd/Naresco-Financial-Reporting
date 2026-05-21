# Naresco Financial Reporting

Multi-company financial reporting system with interactive Budget vs Actual dashboards, line-item drill-downs, and automated data validation.

**Hosting**: Internal server (secure, on-premises)

---

## What This Does

Finance team drops source Excel files → System processes and validates data → Dashboard auto-updates with full drill-down capability per company.

**Key Features:**
- **Multi-Company**: Company selector on load; each company has its own isolated database
- **5 KPI Cards**: Revenue, Contribution Margin %, Gross Profit %, Net Profit %, EBITDA %
- **TDI Split**: Tax/Depreciation/Interest separated from Indirect Costs as its own P&L line
- **Dual-Path Validation**: Excel + JSON must match before deploying
- **Correct YTD Percentages**: Derived from YTD numerator ÷ denominator (not averaged monthly)
- **Correct Variance Convention**: Revenue = Actual − Budget; Costs = Budget − Actual
- **Line-Item Drill-Downs**: Click eye icon on any metric row to see underlying line items
- **Conditional Formatting**: Rules set on parent metrics cascade to child drill-down rows
- **Chart Interactivity**: Hover or click any bar OR its axis label to view values and navigate
- **Always-Fresh Data**: Cache-busting on all JSON fetches — no stale data after pipeline runs
- **Growing Database**: Data accumulates month-over-month per company
- **Secure Hosting**: Data stays on-premises, never leaves internal network

---

## Companies

| ID | Full Name | Source File Pattern |
|---|---|---|
| `mantis` | Mantis | `Mantis*.xlsx` |
| `mudin` | Mudin Al Emerat (MAE) | `Mudin*.xlsx` or `MAE*.xlsx` |

New companies can be added by updating `dashboard/data/companies.json`.

---

## Project Structure

```
Naresco-Financial-Reporting/
│
├── source-files/                       # DROP NEW EXCEL FILES HERE
│   └── README.md
│
├── scripts/                            # Python processing scripts
│   ├── extract_to_excel_database.py    # Extract BVA data to Excel DB (BVA_DATA/BVA_CALC)
│   ├── extract_budget_data.py          # Extract BVA data to JSON
│   ├── extract_revenue_data.py         # Extract Revenue tab line items to DB
│   ├── extract_section_data.py         # Extract cost section line items to DB (config-driven)
│   ├── consolidate_data.py             # Aggregate JSON to YTD totals
│   ├── compare_calculations.py         # Validate Excel vs JSON calculations
│   └── generate_dashboard_from_excel.py # Generate dashboard + detail JSON files
│
├── templates/
│   └── dashboard_template.html         # Static dashboard template (= dashboard/index.html)
│
├── dashboard/                          # Served by web server
│   ├── index.html                      # Dashboard app (copy of template)
│   ├── naresco-logo.png                # Naresco group logo
│   ├── mantis-logo.jpg                 # Mantis company logo
│   └── data/                           # JSON data files (generated, not committed)
│       ├── companies.json              # Company list + logo paths (committed)
│       ├── dashboard-{company}.json    # BVA dashboard data per company
│       ├── calculations-{company}.json # YTD calculations per company
│       ├── raw-data-{company}.json     # Monthly raw data per company
│       ├── revenue-{company}.json      # Revenue line-item detail per company
│       ├── variable-costs-{company}.json
│       ├── fixed-costs-{company}.json
│       ├── indirect-costs-{company}.json  # TDI rows excluded
│       └── tdi-{company}.json             # Interest, Amortisation, Depreciation
│
├── Financial-Data-Database-Mantis.xlsx # Mantis master database
│   ├── BVA_DATA                        # Monthly BVA data
│   ├── BVA_CALC                        # YTD formulas
│   ├── Revenue                         # Contract revenue line items
│   ├── Variable Costs                  # Variable cost line items
│   ├── Fixed Costs                     # Direct staff cost line items
│   └── Indirect Costs                  # Indirect/admin cost line items
│
├── Financial-Data-Database-Mudin.xlsx  # Mudin master database (same structure)
│
├── update-dashboard.bat                # ONE-CLICK UPDATE SCRIPT
├── requirements.txt
└── README.md
```

---

## Monthly Update Process

### For each company, two types of source files go into `source-files/`:

| File type | Example name | Pipeline triggered |
|---|---|---|
| Budget vs Actual | `Mantis Budget vs Actual File Feb 2026.xlsx` | Full BVA pipeline (6 steps) |
| Financial Statement | `Mantis Financial Statement report - FEB 2026.xlsx` | Revenue extraction only |

### Steps:
1. Drop source Excel file(s) into `source-files/`
2. Double-click `update-dashboard.bat`
3. Wait ~1–2 minutes
4. Refresh browser

The script auto-detects the company (from the filename) and the file type, then routes accordingly.

---

## BVA Pipeline (Budget vs Actual files)

```
Source BVA Excel
    ↓
[Step 1] extract_to_excel_database.py   → Financial-Data-Database-{Company}.xlsx
                                            BVA_DATA: append new month rows
                                            BVA_CALC: SUMIF/AVERAGEIF formulas
    ↓
[Step 2] extract_section_data.py        → DB sheets: Variable Costs, Fixed Costs, Indirect Costs
                                            Extracts named line items with B/A/V per month
    ↓
[Step 3] extract_budget_data.py         → dashboard/data/raw-data-{company}.json
    ↓
[Step 4] consolidate_data.py            → dashboard/data/calculations-{company}.json
                                            Absolute metrics: SUM
                                            % metrics: derived from YTD numerator ÷ denominator
                                            EBITDA: Net Surplus + Interest + Depreciation + Amortisation
    ↓
[Step 5] compare_calculations.py        → Validates Excel BVA_CALC vs Python calculations
                                            Stops pipeline if mismatch detected
    ↓
[Step 6] generate_dashboard_from_excel.py → dashboard/data/dashboard-{company}.json
                                            dashboard/data/revenue-{company}.json
                                            dashboard/data/variable-costs-{company}.json
                                            dashboard/data/fixed-costs-{company}.json
                                            dashboard/data/indirect-costs-{company}.json
```

### Financial Statement Pipeline

```
Source Financial Statement Excel
    ↓
extract_revenue_data.py → Revenue sheet in Financial-Data-Database-{Company}.xlsx
                           Detects month columns dynamically
                           Only adds months not already in database
                           YTD Total column always kept rightmost
```

---

## Dashboard Features

### KPI Cards (top row)
| Card | Primary display | Secondary |
|---|---|---|
| Total Revenue | AED value | Variance vs Budget |
| Contribution Margin | % | AED value |
| Gross Profit | % | AED value |
| Net Profit | % | AED value |
| EBITDA | % | AED value |

### Charts
- **Budget vs Actual**: Side-by-side bars for all 7 key metrics
- **Variance Analysis**: Single bar per metric, green = favourable, red = unfavourable
- Clicking any bar scrolls to and highlights the corresponding table row

### Data Table
- All 13 metrics with Budget / Actual / Variance (YTD)
- Eye icon on each row opens the line-item drill-down panel
- Conditional formatting rules (set in Settings) persist per user per browser

### Line-Item Drill-Down Panel
Opens below the data table when eye icon is clicked.

| Metric row | Detail panel shows |
|---|---|
| Total Revenue | Contract revenue by site (actuals only) |
| Total Variable Cost | Material Cost - Projects (B/A/V) |
| Total Staff Cost (Direct) | 12 staff cost lines (B/A/V) |
| Adjusted Indirect Costs | 30 indirect cost lines (B/A/V); TDI components excluded |
| TDI | Interest Expenses, Amortisation, Depreciation (B/A/V) |

**Month selector**: Tick one or more months + Fetch to add columns. YTD Total always visible.

**Zero-variance filter**: Cost sections automatically hide rows with zero variance (rows with no activity). Revenue shows all rows.

**Inherited formatting**: Rules set on a parent metric (e.g. Indirect Costs variance < 0 → red) apply automatically to all child rows in the drill-down.

### P&L Flow (details table order)
```
Total Revenue
− Variable Cost (TVC %)
= Contribution Margin (CM %)
− Fixed Costs Direct (TSCD %)
= Gross Profit (GP %)
− Adjusted Indirect Costs (TFCI %)   ← excludes TDI
= EBITDA (EBITDA %)
− TDI (Tax, Depreciation & Interest)
= Net Surplus / Net Profit (NSD %)
```

### Variance Convention
- **Revenue & profit metrics** (Revenue, CM, GP, EBITDA, Net Profit): Actual − Budget (positive = above budget = good)
- **All cost metrics** (Variable, Staff, Indirect, TDI): Budget − Actual (positive = underspend = good)

This allows management to immediately spot unposted costs: a large positive variance on a cost line may indicate invoices not yet received.

### TDI Split
Interest Expenses, Amortisation and Depreciation are extracted from Indirect Costs and shown as a separate line (TDI — Tax, Depreciation & Interest). This creates a clean EBITDA bridge:
`EBITDA − TDI = Net Surplus`

---

## Metrics Tracked

### BVA Metrics (16 extracted + 4 derived = 20 total)
| Metric | Type | Notes |
|---|---|---|
| Total Revenue | SUM | |
| Total Variable Cost | SUM | |
| TVC % | DERIVED | TVC ÷ Revenue |
| Contribution Margin | SUM | |
| CM % | DERIVED | CM ÷ Revenue |
| Total Staff Cost (Direct) | SUM | |
| TSCD % | DERIVED | TSCD ÷ Revenue |
| Gross Profit / (Loss) | SUM | |
| GP % | DERIVED | GP ÷ Revenue |
| Total Fixed Cost (Indirect) | SUM | Kept for validation; not shown in main table |
| Interest Expenses | SUM | Component of TDI |
| Amortisation | SUM | Component of TDI |
| Depreciation | SUM | Component of TDI |
| TDI | DERIVED | Interest + Amort + Deprec |
| Adjusted Indirect Costs | DERIVED | Total Indirect − TDI; shown as "Indirect Costs" |
| TFCI % | DERIVED | Adjusted Indirect ÷ Revenue |
| Net Surplus / (Deflect) | SUM | |
| NSD % | DERIVED | NSD ÷ Revenue |
| EBITDA | DERIVED | Net Surplus + Interest + Amort + Deprec |
| EBITDA % | DERIVED | EBITDA ÷ Revenue |

---

## Adding a New Section Drill-Down

To add a new cost section (e.g. "Marketing Costs"):

**1. `scripts/extract_section_data.py`** — add to `SECTIONS_CONFIG`:
```python
"Marketing Costs": {
    "metrics": ["Row Label 1", "Row Label 2", ...],
    "total_label": "TOTAL MARKETING COSTS",
    "source_sheet": "Detail Budget"
},
```

**2. `scripts/generate_dashboard_from_excel.py`** — add to `SECTION_SHEETS`:
```python
"Marketing Costs": "marketing-costs",
```

**3. `templates/dashboard_template.html`** — add to `DETAIL_METRICS`:
```javascript
'Metric Row Key': {type: 'section', slug: 'marketing-costs', label: 'Marketing Costs', filterZeroVariance: true},
```

**4. `scripts/generate_dashboard_from_excel.py`** — if any rows should be excluded from the JSON (e.g. they belong to another section), add to the `exclude` set in `SECTION_SHEETS`:
```python
"My Section": {"slug": "my-section", "exclude": {"Row to hide"}},
```

Copy template to `dashboard/index.html`, run the pipeline. Done.

---

## Adding a New Company

1. Add entry to `dashboard/data/companies.json`:
```json
{"id": "newco", "name": "New Company Name", "logo": "newco-logo.png"}
```
2. Place logo file in `dashboard/`
3. Drop BVA and Financial Statement source files (with company name in filename) into `source-files/`
4. Run `update-dashboard.bat`

---

## Tech Stack

- **Python 3.8+**: openpyxl, pandas
- **JavaScript**: Chart.js 4.4 for visualisations
- **Storage**: Per-company Excel databases + generated JSON
- **Hosting**: Internal web server (IIS recommended for production)
- **Version Control**: Git / GitHub

---

## Local Testing

```bash
cd dashboard
python -m http.server 8000
# Open http://localhost:8000
```

The dashboard uses `fetch()` to load JSON data, so a local server is required (cannot open index.html directly as a file).

---

## Troubleshooting

**"No Excel files found"** — Ensure `.xlsx` file is in `source-files/` folder

**"Validation failed"** — Check BVA_DATA and BVA_CALC sheets in the company database; verify source file structure matches expected format

**"Could not load data / No detail data found"** — Run the full update pipeline; JSON files are not committed to git and must be generated locally

**"Permission denied on database"** — Close the Excel database file before running the pipeline

**"Dashboard not updating"** — Hard refresh browser (Ctrl+F5); ensure the pipeline completed successfully

**"Admin settings disappeared"** — Settings are stored per browser. To transfer: copy `localStorage.getItem('formattingRules')` from browser console and paste on new browser with `localStorage.setItem(...)`

---

*Built for Naresco Finance | Multi-company Budget vs Actual reporting*
