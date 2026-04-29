# Naresco Financial Reporting - Project Documentation

## Project Overview

**Purpose**: Automated financial reporting infrastructure for Budget vs Actual analysis at Naresco Finance.

**Goal**: Finance team drops Excel files into a folder → System automatically extracts data → Generates interactive web dashboards → Deploys to Vercel.

**Status**: Phase 1 (Budget vs Actual Dashboard) - **COMPLETE & DEPLOYED**
- Live Dashboard: https://naresco-financial-reporting.vercel.app
- GitHub Repo: https://github.com/charmenlondon-cmd/Naresco-Financial-Reporting

---

## Architecture

```
Excel Files (Source) 
    ↓
[1] Excel → Values-Only Copy (removes formulas, external links)
    ↓
[2] Data Extraction (Python) → Extract 13 metrics → raw-data.json
    ↓
[3] Data Consolidation (Python) → Aggregate YTD totals → calculations.json
    ↓
[4] Dashboard Generation (Python) → Inject data into HTML template → index.html
    ↓
[5] Git Commit & Push → Vercel Auto-Deploy
    ↓
Live Dashboard (Web)
```

### Tech Stack

- **Data Processing**: Python 3.8+ (pandas, openpyxl)
- **Automation**: n8n (planned)
- **Dashboard**: HTML, CSS, JavaScript, Chart.js
- **Hosting**: Vercel (static site)
- **Data Storage**: JSON files (Google Drive planned for POC)
- **Version Control**: Git + GitHub

---

## Key Design Decisions

### 1. Name-Based Metric Extraction (Not Row Numbers)
**Why**: Excel file structure can change (rows added/removed)
**How**: Search for exact metric names in column A (e.g., "Total Revenue", "Total Variable Cost")
**Percentage Handling**: Always in row directly below parent metric (e.g., "TVC %" below "Total Variable Cost")

### 2. Month Detection Algorithm
**Method**: Count "Budget" and "Total Budget" headers in row 11
**Formula**: Number of months = Count of columns with "Budget" or "Total Budget" header
**Why**: Robust against file structure changes, includes cumulative totals as final month
**Note**: Handles "Total Budget" column as Month 3 (cumulative YTD data)

### 3. Multi-Month Extraction (Default)
**Behavior**: Extracts ALL months from file by default
**Flag**: `--last-month-only` to extract only the most recent month
**Why**: Accumulates historical data for YTD analysis

### 4. Aggregation Logic
**Absolute Values** (Revenue, Costs, Margins): **SUM** across months
**Percentage Values** (TVC %, CM %, etc.): **AVERAGE** across months
**Why**: Matches Excel SUMIF/AVERAGEIF behavior from sample dashboard

### 5. Currency: AED (UAE Dirham)
**Format**: `AED 1,000,000`
**Locale**: `en-AE`
**Why**: Company operates in UAE

### 6. Typography: Inter Font
**Why**: Modern, rounded, ultra-readable, used by top tech companies
**Fallback**: -apple-system, BlinkMacSystemFont, Segoe UI

### 7. Dashboard Philosophy
**Focus**: Executive summary, variance analysis, detailed table
**Removed**: Confusing trend charts (user feedback)
**Kept**: Variance analysis (most useful), KPI cards, admin controls

---

## File Structure

```
naresco-financial-reporting/
├── scripts/                          # Data processing scripts
│   ├── extract_budget_data.py       # Extracts metrics from Excel → raw-data.json
│   ├── consolidate_data.py          # Aggregates monthly data → calculations.json
│   └── generate_dashboard.py        # Generates HTML dashboard from data
│
├── templates/                        # Dashboard HTML template
│   └── dashboard_template.html      # Template with {{DASHBOARD_DATA}} placeholder
│
├── config/                           # Configuration files
│   └── workflow_config.example.json # Example configuration
│
├── dashboard/                        # Generated dashboard (deployed to Vercel)
│   ├── index.html                   # Generated dashboard (self-contained)
│   ├── assets/                      # CSS, JS (currently inline in template)
│   └── data/                        # JSON data files (gitignored)
│       ├── raw-data.json           # Monthly metrics (all months)
│       └── calculations.json       # YTD aggregated totals
│
├── n8n-workflows/                    # Workflow automation (planned)
│
├── .gitignore                        # Excludes sensitive data
├── README.md                         # Project overview
├── CLAUDE.md                         # This file - project documentation
├── requirements.txt                  # Python dependencies
├── vercel.json                       # Vercel deployment config
└── .git/                            # Git repository
```

---

## Data Flow & File Formats

### Input: Excel File
**File Name Pattern**: `*Budget vs Actual File*.xlsx` (e.g., "Mudin Budget vs Actual File 2026.xlsx")
**Sheet**: "Detail Budget"
**Structure**:
- Row 11: Headers (Budget, Actual, Variance repeated for each month)
- Column A: Metric names
- Columns 2-4: January (Budget, Actual, Variance)
- Columns 5-7: February (Budget, Actual, Variance)
- etc.

### Intermediate: raw-data.json
**Format**: Array of monthly metric records
```json
[
  {
    "Month": "January",
    "DataPoint": "Total Revenue",
    "DashboardName": "Total Revnue",
    "Budget": 9000000.0,
    "Actual": 4000000.0,
    "Variance": -5000000.0
  },
  ...
]
```

### Intermediate: calculations.json
**Format**: Array of YTD aggregated metrics
```json
[
  {
    "DataPoint": "Total Revenue",
    "DashboardName": "Total Revnue",
    "YTD_BUDGET": 17000000.0,
    "YTD_ACTUAL": 10000000.0,
    "YTD_VARIANCE": -7000000.0,
    "AggregationMethod": "SUM"
  },
  ...
]
```

### Output: index.html
**Format**: Self-contained HTML with embedded JSON data
**Data Injection**: `{{DASHBOARD_DATA}}` placeholder replaced with JSON
**Dependencies**: Chart.js (CDN), Inter font (Google Fonts)

---

## Metrics Tracked (13 Total)

| # | Metric Name | Database Name | Dashboard Display | Type | Aggregation |
|---|-------------|---------------|-------------------|------|-------------|
| 1 | Total Revenue | Total Revenue | Total Revnue | Absolute | SUM |
| 2 | Total Variable Cost | Total Variable Cost | Variable Cost | Absolute | SUM |
| 3 | TVC % | TVC % | Variable Cost % | Percentage | AVERAGE |
| 4 | Contribution Margin | Contribution Margin | Contribution Margin | Absolute | SUM |
| 5 | CM % | CM % | Contribution Margin % | Percentage | AVERAGE |
| 6 | Total Staff Cost (Direct) | Total Staff Cost (Direct) | Fixed Costs (Direct) | Absolute | SUM |
| 7 | TSCD % | TSCD % | Fixed Costs (Direct) % | Percentage | AVERAGE |
| 8 | Gross Profit / (Loss) | Gross Profit / (Loss) | Gross Profit / Loss* | Absolute | SUM |
| 9 | GP % | GP % | Gross Profit / Loss % | Percentage | AVERAGE |
| 10 | Total Fixed Cost (Indirect) | Total Fixed Cost (Indirect) | Fixed Costs (Indirect) | Absolute | SUM |
| 11 | TFCI % | TFCI % | Fixed Costs (Indirect) % | Percentage | AVERAGE |
| 12 | Net Surplus / (Deflect) | Net Surplus / (Deflect) | Net Profit / Loss* | Absolute | SUM |
| 13 | NSD % | NSD % | Net Profit / Loss % | Percentage | AVERAGE |

**Note**: *Labels with asterisk are **dynamic** - change based on value:
- Positive value → "Gross Profit" / "Net Profit"
- Negative value → "Gross Loss" / "Net Loss"

---

## Scripts Usage

### 1. Extract Data from Excel
```bash
python scripts/extract_budget_data.py \
  --input "Mudin Budget vs Actual File 2026.xlsx" \
  --output "dashboard/data/raw-data.json" \
  --sheet "Detail Budget"

# Extract only last month:
python scripts/extract_budget_data.py --input "file.xlsx" --last-month-only
```

### 2. Consolidate Monthly Data
```bash
python scripts/consolidate_data.py \
  --input "dashboard/data/raw-data.json" \
  --output "dashboard/data/calculations.json"
```

### 3. Generate Dashboard
```bash
python scripts/generate_dashboard.py \
  --raw-data "dashboard/data/raw-data.json" \
  --calculations "dashboard/data/calculations.json" \
  --template "templates/dashboard_template.html" \
  --output "dashboard/index.html"
```

### 4. Complete Pipeline (All 3 Steps)
```bash
# Process new Excel file
python scripts/extract_budget_data.py --input "NewFile.xlsx"
python scripts/consolidate_data.py
python scripts/generate_dashboard.py

# Commit and deploy
git add dashboard/index.html
git commit -m "Update dashboard with [Month] data"
git push  # Vercel auto-deploys
```

---

## Dashboard Features

### Executive KPI Cards
- **Total Revenue**: Budget vs Actual with variance
- **Contribution Margin**: Budget vs Actual with variance
- **Gross Profit/Loss**: Dynamic label based on ACTUAL value
- **Net Profit/Loss**: Dynamic label based on ACTUAL value

### Budget vs Actual Comparison Chart (NEW)
- **Type**: Grouped bar chart
- **Position**: Below KPI cards, above Variance Analysis
- **Data**: Budget (blue bars) vs Actual (green bars) for 7 key metrics
- **Labels**: Based on ACTUAL values (Gross Profit/Loss, Net Profit/Loss)
- **Purpose**: Visual comparison of absolute performance vs plan
- **Style**: Rounded corners, smooth animations, side-by-side comparison

### Variance Analysis Chart
- **Type**: Horizontal bar chart
- **Header**: Prominent "Variance Analysis" (bold, blue, 24px)
- **Data**: YTD variance for 7 key metrics
- **Labels**: Based on VARIANCE direction (Gross Loss Variance when negative)
- **Colors**: 
  - Cost metrics: Negative variance = green (under budget), Positive = red (over budget)
  - Revenue/Profit metrics: Positive variance = green (over budget), Negative = red (under budget)
- **Purpose**: Shows performance gaps and budget deviations
- **Style**: Rounded corners, smooth animations

### Data Table
- **Columns**: Metric, Budget, Actual, Variance ($)
- **Rows**: All 13 metrics in logical order
- **Formatting**: Conditional formatting via admin controls (both Actual and Variance columns)
- **Font**: Poppins (modern, rounded, readable)
- **Dynamic Labels**: Profit/Loss labels change based on ACTUAL values

### Admin Control Panel
- **Access**: "Admin Controls" button in header
- **Features**:
  - Configure up to 3 formatting rules per metric
  - **Column Selector**: Choose Actual or Variance column per rule
  - Conditions: >, <, >=, <=, =
  - Actions: Text color or background highlight
  - Color picker for custom colors
  - Rules saved in browser localStorage
  - **Rule Stacking**: One color + one background rule can both apply per column
- **Rule Priority**: First matching rule wins within each property type (color/background)

---

## Deployment (Vercel)

### Configuration
**File**: `vercel.json`
```json
{
  "buildCommand": "echo 'Using pre-built dashboard'",
  "outputDirectory": "dashboard",
  "framework": null
}
```

### Workflow
1. Generate dashboard locally (Python scripts)
2. Commit `dashboard/index.html` to git
3. Push to GitHub main branch
4. Vercel auto-deploys in ~30 seconds

### Important Notes
- Dashboard is **static HTML** (no Python server needed)
- Data is **embedded** in HTML (self-contained)
- No build process on Vercel (pre-generated locally)
- JSON data files are gitignored (only generated HTML is committed)

---

## Important Implementation Details

### Month Detection Logic
**Row 11** contains column headers: "Budget", "Actual", "Variance" repeated for each month.

**Algorithm**:
```python
month_count = count_of_columns_with_"Budget"_header_in_row_11
month_number = month_count  # 1 = January, 2 = February, etc.
```

**Why not count total columns?**
- File may have additional columns with non-month data
- Headers in row 11 are reliable indicators of actual month columns

### Percentage Metric Naming Convention
**Pattern**: First letter of each word + " %"

Examples:
- **T**otal **V**ariable **C**ost → TVC %
- **C**ontribution **M**argin → CM %
- **T**otal **S**taff **C**ost (**D**irect) → TSCD %
- **G**ross **P**rofit → GP %
- **T**otal **F**ixed **C**ost (**I**ndirect) → TFCI %
- **N**et **S**urplus / (**D**eflect) → NSD %

**Location in Source File**: Always in row directly below parent metric, with "%" in column A

### Error Handling
**Excel Errors**: #REF!, #VALUE!, #N/A → Converted to 0.0
**Missing Metrics**: Logged as warning, skipped (doesn't break extraction)
**Invalid Values**: Non-numeric values → 0.0

---

## Git Workflow

### .gitignore Strategy
**Excluded** (sensitive/generated):
- `*.xlsx`, `*.xls` - Source Excel files
- `dashboard/data/*.json` - Data files
- `config/workflow_config.json` - Config with credentials
- `*.key`, `*.pem` - Credentials

**Included** (needed for deployment):
- `dashboard/index.html` - Generated dashboard (self-contained)
- `templates/` - Template files
- `scripts/` - Python scripts
- `vercel.json` - Deployment config

### Commit Message Format
```
<verb>: <description>

- Detailed change 1
- Detailed change 2

Co-Authored-By: Claude Sonnet 4.5 <noreply@anthropic.com>
```

Examples:
- `Add data extraction script`
- `Fix: Remove annoying alert on page load`
- `Enhance dashboard: AED currency + premium chart styling`

---

## Phase 1 Completion Status

### ✅ COMPLETE
1. **Data Extraction** - Extract 13 metrics from Excel files
2. **Data Consolidation** - Aggregate monthly data to YTD
3. **Dashboard Generation** - Create interactive HTML dashboard
4. **Dashboard Design** - Professional, modern UI with Inter font
5. **Chart Visualization** - Variance analysis with gradients
6. **Admin Controls** - Conditional formatting panel
7. **Vercel Deployment** - Live at naresco-financial-reporting.vercel.app
8. **Currency Support** - AED formatting
9. **Dynamic Labels** - Profit/Loss changes based on value
10. **GitHub Integration** - Push-to-deploy workflow

### 🔄 PLANNED (Phase 2+)
1. **n8n Workflow Automation**
   - Monitor folder for new Excel files
   - Auto-run extraction → consolidation → generation
   - Auto-commit and push to trigger deployment

2. **Google Drive Integration**
   - Store JSON data files
   - API integration for file upload/download

3. **Interactive Features** (Next Priority)
   - Click interactions on charts/cards
   - Drill-down capabilities
   - Month selector/filter
   - Export to Excel/PDF
   - Search/filter in data table

4. **Additional Dashboards**
   - P&L Dashboard
   - Cash Flow Dashboard
   - Financial KPIs Dashboard
   - Multi-company support

5. **SharePoint Integration**
   - Replace file folder monitoring with SharePoint API
   - Production-ready file source

---

## Configuration Files

### workflow_config.example.json
**Purpose**: Template for workflow configuration
**Contains**:
- Company name
- File patterns
- Metric names and display labels
- Google Drive settings (when implemented)

**Usage**: Copy to `workflow_config.json` and customize

---

## Testing

### Current Test Data
**File**: `Mudin Budget vs Actual File 2026_values.xlsx`
**Months**: January, February (2 months)
**Metrics**: All 13 metrics with Budget, Actual, Variance

### Test Workflow
```bash
# 1. Extract data
python scripts/extract_budget_data.py \
  --input "Mudin Budget vs Actual File 2026_values.xlsx"

# 2. Consolidate
python scripts/consolidate_data.py

# 3. Generate dashboard
python scripts/generate_dashboard.py

# 4. Open in browser
# dashboard/index.html
```

### Verification Checklist
- [ ] All 13 metrics extracted
- [ ] Both months (Jan, Feb) present in raw-data.json
- [ ] YTD totals correct in calculations.json
- [ ] Dashboard displays without errors
- [ ] Charts render correctly
- [ ] Currency shows as AED
- [ ] Dynamic labels work (Profit/Loss)
- [ ] Admin controls functional
- [ ] No alert popup on page load

---

## Known Issues & Solutions

### Issue: Stream idle timeout errors
**Cause**: Large command outputs or network latency
**Solution**: Keep commands short, break into smaller chunks

### Issue: Unicode characters in console output
**Cause**: Windows console encoding (cp1252)
**Solution**: Use ASCII characters ([OK], [ERROR]) instead of emoji (✓, ❌)

### Issue: Excel #REF! errors in values-only file
**Cause**: Original file had formula reference errors
**Solution**: Error handling converts #REF!, #VALUE!, etc. to 0.0

---

## Future Enhancements (Ideas)

### User Experience
- [ ] Dark mode toggle
- [ ] Mobile-responsive design improvements
- [ ] Loading animations
- [ ] Export capabilities (PDF, Excel, CSV)
- [ ] Print-friendly stylesheet

### Data & Analytics
- [ ] Month-over-month comparison
- [ ] Year-over-year comparison
- [ ] Trend forecasting
- [ ] Anomaly detection
- [ ] Custom date range selection

### Automation
- [ ] Email notifications when dashboard updates
- [ ] Slack integration for alerts
- [ ] Scheduled reports
- [ ] Automatic data backup

### Multi-Company
- [ ] Support multiple companies
- [ ] Company selector dropdown
- [ ] Consolidated multi-company view
- [ ] Comparison across companies

---

## Team & Contact

**Project Owner**: Naresco Finance
**Primary User**: Finance Team
**Developer**: Claude Sonnet 4.5 (Anthropic)
**GitHub**: https://github.com/charmenlondon-cmd/Naresco-Financial-Reporting
**Live Dashboard**: https://naresco-financial-reporting.vercel.app

---

## Documentation Updates

**Last Updated**: 2026-04-29 (Evening Session)
**Version**: 1.1.0 (Phase 1 Complete + Enhancements)
**Next Review**: After Phase 2 (Interactivity Features)

### Recent Updates (2026-04-29 Evening)

**3-Month Data Extraction:**
- Fixed month detection to recognize "Total Budget" column as Month 3
- Now correctly extracts January, February, and March (cumulative) data
- 39 records total (3 months × 13 metrics)
- Resolved issue with hidden junk columns in source Excel file

**Dashboard Enhancements:**
- Added "Budget vs Actual Comparison" chart (grouped bars, blue vs green)
- Positioned above Variance Analysis for better workflow
- Updated Variance Analysis with prominent header and variance-based labels
- Removed unauthorized Variance % column from data table

**Admin Control Improvements:**
- Added column selector (Actual or Variance) for each formatting rule
- Rules can now format both Actual and Variance columns independently
- Fixed rule stacking: one color + one background rule can both apply
- First matching rule wins within each property type to prevent conflicts

**Bar Chart Logic Fixes:**
- Budget vs Actual chart: Labels based on ACTUAL values (Gross Profit/Loss)
- Variance chart: Labels based on VARIANCE direction (Gross Loss Variance when negative)
- Variance colors: Based on performance (cost metrics inverted)
- Bars correctly extend below zero for negative variances

**Technical Improvements:**
- Excel database extraction script created (extract_to_excel_database.py)
- Ready for accumulating data over time in Excel format
- Self-contained dashboard with embedded JSON data

---

## Quick Start Guide

### For New Team Members

1. **Clone Repository**
   ```bash
   git clone https://github.com/charmenlondon-cmd/Naresco-Financial-Reporting.git
   cd Naresco-Financial-Reporting
   ```

2. **Install Python Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Process a File**
   ```bash
   python scripts/extract_budget_data.py --input "YourFile.xlsx"
   python scripts/consolidate_data.py
   python scripts/generate_dashboard.py
   ```

4. **View Dashboard Locally**
   Open `dashboard/index.html` in your browser

5. **Deploy to Vercel**
   ```bash
   git add dashboard/index.html
   git commit -m "Update dashboard"
   git push
   ```

### For Finance Team

1. **Drop Excel file** in designated folder (future: SharePoint)
2. **Wait for automation** (future: n8n handles this)
3. **View dashboard** at https://naresco-financial-reporting.vercel.app
4. **Use admin controls** to customize formatting if needed

---

**End of Project Documentation**
