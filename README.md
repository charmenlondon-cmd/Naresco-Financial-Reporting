# Naresco Financial Reporting

Automated financial reporting system with dual-path validation for Budget vs Actual dashboards.

**Live Dashboard**: https://naresco-financial-reporting.vercel.app

---

## 🎯 What This Does

Finance team drops a new Excel file → System processes and validates data → Dashboard auto-updates and deploys

**Key Features:**
- ✅ **One-Click Updates**: Drop file + double-click = Done
- ✅ **Dual-Path Validation**: Excel + JSON must match before deploying
- ✅ **Growing Database**: Data accumulates month-over-month
- ✅ **Interactive Dashboard**: Click KPI cards to jump to details
- ✅ **Auto-Deployment**: Vercel deploys in ~30 seconds

---

## 📁 Project Structure

```
Naresco-Financial-Reporting/
│
├── source-files/                    # DROP NEW EXCEL FILES HERE
│   └── README.md
│
├── scripts/                         # Python processing scripts
│   ├── extract_to_excel_database.py    # PATH A: Extract to Excel
│   ├── extract_budget_data.py          # PATH B: Extract to JSON
│   ├── consolidate_data.py             # PATH B: Consolidate JSON
│   ├── compare_calculations.py         # Validate Excel vs JSON
│   └── generate_dashboard_from_excel.py # Generate from validated Excel
│
├── templates/
│   └── dashboard_template.html      # Dashboard HTML template
│
├── dashboard/                       # Generated dashboard (deployed)
│   ├── index.html                   # Generated dashboard
│   ├── naresco-logo.png            # Company logo
│   └── data/                       # JSON files (validation path)
│
├── Financial-Data-Database.xlsx     # MASTER DATABASE
│   ├── BVA_DATA                    # Monthly accumulated data
│   └── BVA_CALC                    # YTD formulas
│
├── update-dashboard.bat             # ONE-CLICK UPDATE SCRIPT
├── requirements.txt
└── README.md                       # This file
```

---

## 🚀 Quick Start (For Finance Team)

### **Monthly Update Process:**

1. **Drop File**: Save new Excel file to `source-files/` folder
2. **Run Script**: Double-click `update-dashboard.bat`
3. **Wait**: ~1-2 minutes for processing
4. **Done**: Dashboard live at https://naresco-financial-reporting.vercel.app

**That's it!** The script handles everything automatically.

---

## 🏗️ How It Works (Dual-Path Validation)

```
Source Excel File (New Month)
    ↓
    ├─→ [PATH A: Excel Database - MASTER]
    │   ├─ Extract to Financial-Data-Database.xlsx
    │   ├─ BVA_DATA: Append new month
    │   └─ BVA_CALC: Formulas auto-calculate YTD
    │
    └─→ [PATH B: JSON - VALIDATION]
        ├─ Extract to raw-data.json
        └─ Consolidate to calculations.json
    
    ↓
[VALIDATION]
Compare Excel vs JSON
    ├─ ✅ Match → Proceed to dashboard
    └─ ❌ Mismatch → STOP! Alert user
    
    ↓
[DASHBOARD]
Generate from Excel (validated source)
    ↓
Git push → Vercel auto-deploy
```

**Why Dual-Path?**
- Excel = Master (auditable, human-readable)
- JSON = Independent validation
- If they don't match = Something is wrong, don't deploy!

---

## 📊 Data Flow

### **Month 1 (February):**
```
Feb Excel → update-dashboard.bat
    ↓
BVA_DATA: 13 rows (Feb only)
BVA_CALC: YTD for 1 month
Dashboard: Shows Feb data
```

### **Month 2 (March):**
```
Mar Excel → update-dashboard.bat
    ↓
BVA_DATA: 26 rows (Feb + Mar)  ← Accumulated!
BVA_CALC: YTD for 2 months
Dashboard: Shows Feb + Mar
```

### **Month 3 (April):**
```
Apr Excel → update-dashboard.bat
    ↓
BVA_DATA: 39 rows (Feb + Mar + Apr)  ← Growing!
BVA_CALC: YTD for 3 months
Dashboard: Shows Feb + Mar + Apr
```

**Data never gets lost - it accumulates over time!**

---

## 🔧 Technical Details

### **Metrics Tracked (13 total):**
- Total Revenue
- Variable Cost + %
- Contribution Margin + %
- Fixed Costs (Direct) + %
- Gross Profit/Loss + %
- Fixed Costs (Indirect) + %
- Net Profit/Loss + %

### **Dashboard Features:**
- **Interactive KPI Cards**: Click to jump to row in table
- **Budget vs Actual Chart**: Side-by-side comparison
- **Variance Analysis Chart**: Performance gaps visualization
- **Data Table**: All metrics with conditional formatting
- **Admin Controls**: Custom formatting rules (saved in browser)

### **Tech Stack:**
- **Python 3.8+**: pandas, openpyxl
- **Excel**: Master database with formulas
- **JavaScript**: Chart.js for visualizations
- **Hosting**: Vercel (auto-deploy)
- **Version Control**: Git + GitHub

---

## 🛠️ Installation (For Developers)

### **Prerequisites:**
- Python 3.8+
- Git

### **Setup:**

```bash
# Clone repository
git clone https://github.com/charmenlondon-cmd/Naresco-Financial-Reporting.git
cd Naresco-Financial-Reporting

# Install dependencies
pip install -r requirements.txt

# Ready to use!
```

---

## 📝 Manual Commands (Advanced)

If you prefer manual control instead of using the batch script:

```bash
# Step 1: Extract to Excel database
python scripts/extract_to_excel_database.py --input "source-files/YourFile.xlsx"

# Step 2: Extract to JSON (validation)
python scripts/extract_budget_data.py --input "source-files/YourFile.xlsx"
python scripts/consolidate_data.py

# Step 3: Validate
python scripts/compare_calculations.py

# Step 4: Generate dashboard
python scripts/generate_dashboard_from_excel.py

# Step 5: Deploy
git add dashboard/index.html
git commit -m "Update dashboard"
git push
```

---

## ✅ Validation Example

### **When Excel and JSON Match (Good!):**
```
✅ VALIDATION PASSED!
All calculations verified and matched!

  ✓ Total Revenue                Excel=    20,000,000.00  JSON=    20,000,000.00  ✅
  ✓ Total Variable Cost          Excel=    16,000,000.00  JSON=    16,000,000.00  ✅
  ✓ Contribution Margin          Excel=     4,000,000.00  JSON=     4,000,000.00  ✅
  ... (all 13 metrics checked)

Proceeding with dashboard generation...
```

### **When They Don't Match (Problem!):**
```
❌ VALIDATION FAILED!

  ❌ Total Variable Cost          MISMATCH:
     Actual:   Excel=16,000,000  JSON=16,500,000  Diff=500,000

Process STOPPED. Fix issues before deploying.
```

---

## 🔄 Workflow Automation (Planned - Phase 2)

Currently manual (double-click batch file).

**Future: n8n Integration**
- Monitor `source-files/` folder automatically
- Trigger batch script when new file detected
- Send notifications on completion/errors

---

## 📈 Roadmap

**Phase 1 (Complete)**: Budget vs Actual Dashboard ✅
- Dual-path validation system
- Growing Excel database
- Interactive web dashboard
- One-click updates

**Phase 2 (In Progress)**: P&L Dashboard 🔄
- Profit & Loss analysis
- Year-over-year comparison
- Multi-period trending

**Phase 3 (Planned)**: Additional Reports
- Balance Sheet
- Cash Flow
- Multi-company support

---

## 🐛 Troubleshooting

**"No Excel files found"**
- Ensure .xlsx file is in `source-files/` folder

**"Validation failed"**
- Check `Financial-Data-Database.xlsx` (BVA_DATA and BVA_CALC sheets)
- Verify source Excel file structure matches expected format
- Review validation output for specific mismatches

**"Git push failed"**
- Authenticate with GitHub
- Run `git push` manually to set up credentials

---

## 📞 Support

For issues or questions, contact the Naresco Finance team.

---

**Built for Naresco Finance | Powered by dual-path validation**
