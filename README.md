# Naresco Financial Reporting

Automated financial reporting infrastructure for Budget vs Actual dashboards.

## Overview

This system automates the extraction, processing, and visualization of financial data from Excel files into interactive web-based dashboards.

### Key Features

- 🔄 **Automated Data Pipeline**: Monitors for new financial files and processes them automatically
- 📊 **Dynamic Dashboards**: Interactive web-based dashboards that grow month-over-month
- 🎯 **Budget vs Actual Analysis**: Track variances across revenue, costs, and profitability metrics
- 🎨 **Custom Formatting Controls**: Admin panel for configuring conditional formatting rules
- 📈 **Year-over-Year Tracking**: Accumulates historical data with automatic year-end reset

## Architecture

```
Source Files (Excel) → Values-Only Conversion → Data Extraction → 
JSON Storage → Dashboard Generation → Web Deployment (Vercel)
```

**Tech Stack:**
- Python (pandas, openpyxl) for data processing
- n8n for workflow automation
- HTML/JavaScript (Chart.js) for dashboards
- Google Drive for data storage (POC)
- Vercel for dashboard hosting

## Project Structure

```
naresco-financial-reporting/
├── scripts/              # Python data processing scripts
├── templates/            # Dashboard HTML templates
├── config/              # Configuration files
├── n8n-workflows/       # Workflow automation definitions
└── dashboard/           # Generated dashboard files
    ├── assets/          # CSS, JavaScript, images
    └── data/            # Generated JSON data
```

## Getting Started

### Prerequisites

- Python 3.8+
- n8n (local or cloud instance)
- Google Drive API access (for POC)
- Git and GitHub CLI

### Installation

```bash
# Clone repository
git clone https://github.com/charmenlondon-cmd/Naresco-Financial-Reporting.git
cd Naresco-Financial-Reporting

# Install Python dependencies
pip install -r requirements.txt

# Configure settings
cp config/workflow_config.example.json config/workflow_config.json
# Edit config/workflow_config.json with your settings
```

## Workflow

1. **File Drop**: Finance team drops new Excel file in monitored folder
2. **Conversion**: System creates values-only copy (removes formulas)
3. **Extraction**: Extracts 13 key metrics (Revenue, Costs, Margins, Profits)
4. **Storage**: Appends to historical JSON database
5. **Calculation**: Aggregates YTD totals and percentages
6. **Dashboard**: Generates updated HTML dashboard
7. **Deployment**: Auto-deploys to Vercel

## Metrics Tracked

- Total Revenue
- Variable Costs (and %)
- Contribution Margin (and %)
- Fixed Costs - Direct (and %)
- Gross Profit/Loss (and %)
- Fixed Costs - Indirect (and %)
- Net Profit/Loss (and %)

## Dashboard Features

- **Executive KPI Cards**: High-level metrics at a glance
- **Trend Visualizations**: Budget vs Actual over time
- **Variance Analysis**: Detailed breakdown by category
- **Conditional Formatting**: Customizable color-coding rules
- **Dynamic Labels**: "Profit" vs "Loss" based on values
- **Growing Time Series**: Accumulates data month-over-month

## Configuration

See `config/workflow_config.json` for:
- File naming patterns
- Metric extraction rules
- Google Drive folder paths
- Dashboard customization options

## Development

### Running Data Extraction

```bash
python scripts/extract_budget_data.py --input "path/to/file.xlsx" --output "data/raw-data.json"
```

### Generating Dashboard

```bash
python scripts/generate_dashboard.py --data "data/" --output "dashboard/index.html"
```

### Testing Workflow

```bash
# Process test file
python scripts/extract_budget_data.py --input "test/sample.xlsx"
python scripts/consolidate_data.py
python scripts/generate_dashboard.py
```

## Deployment

Dashboard is automatically deployed to Vercel on push to `main` branch.

**Manual deployment:**
```bash
vercel --prod
```

## Roadmap

**Phase 1 (Current)**: Budget vs Actual Dashboard
- ✅ Data extraction pipeline
- ✅ Dashboard generation
- 🔄 n8n workflow integration
- 🔄 Vercel deployment

**Phase 2**: P&L Dashboard
- Monthly P&L statements
- Revenue and expense breakdowns
- Trend analysis

**Phase 3**: Additional Reports
- Cash Flow forecasting
- Financial KPIs dashboard
- Multi-company support

## Contributing

This is a private project for Naresco Finance. For questions or issues, contact the development team.

## License

Proprietary - All rights reserved

---

**Built with ❤️ for Naresco Finance**
