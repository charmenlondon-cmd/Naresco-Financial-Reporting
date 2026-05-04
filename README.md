# Naresco Financial Reporting

Automated financial reporting system with dual-path validation for Budget vs Actual dashboards.

**Hosting**: Internal server (secure, on-premises)

---

## 🎯 What This Does

Finance team drops a new Excel file → System processes and validates data → Dashboard auto-updates

**Key Features:**
- ✅ **One-Click Updates**: Drop file + double-click = Done
- ✅ **Dual-Path Validation**: Excel + JSON must match before deploying
- ✅ **Growing Database**: Data accumulates month-over-month
- ✅ **Interactive Dashboard**: Click KPI cards to jump to details
- ✅ **Shorthand Formatting**: Fortune 500-style numbers (AED 1.28M vs AED 1,279,576)
- ✅ **Personalized Views**: Each user customizes their own dashboard settings
- ✅ **Secure Hosting**: Data stays on-premises, never leaves internal network

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
4. **Done**: Refresh browser to see updated dashboard

**That's it!** The script handles everything automatically.

### **Accessing the Dashboard:**

**Internal Server**: Navigate to your internal dashboard URL (e.g., `http://naresco-server/financial-dashboard`)

**Local Testing**: Run `python -m http.server 8000` in the `dashboard/` folder, then open `http://localhost:8000`

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
Serve via internal web server
```

**Why Dual-Path?**
- Excel = Master (auditable, human-readable)
- JSON = Independent validation
- If they don't match = Something is wrong, don't deploy!

**Security:**
- Data hosted internally on secure on-premises server
- No external cloud dependencies
- Full control over access and permissions

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
- **Shorthand Number Formatting**: Fortune 500-style (AED 1.28M instead of AED 1,279,576)
- **Budget vs Actual Chart**: Side-by-side comparison
- **Variance Analysis Chart**: Performance gaps visualization
- **Data Table**: All metrics with full precision and conditional formatting
- **Admin Controls**: Custom formatting rules (saved per user in browser)
- **Personalized Views**: Each user's browser stores their own settings

### **Number Formatting:**
- **KPI Cards**: Shorthand format (up to 2 decimals, trailing zeros removed)
  - Millions: AED 1.28M
  - Thousands: AED 15.5K
  - Small numbers: AED 250
- **Detail Table**: Full precision (AED 1,279,576)
- **Design Philosophy**: "Overview first, details on demand"

### **Tech Stack:**
- **Python 3.8+**: pandas, openpyxl
- **Excel**: Master database with formulas
- **JavaScript**: Chart.js for visualizations
- **Hosting**: Internal web server (IIS, Apache, nginx, or Python http.server)
- **Version Control**: Git (optional)

---

## 🛠️ Installation & Deployment

### **Prerequisites:**
- Python 3.8+
- Web server (IIS, Apache, nginx, or Python's built-in server)

### **Setup:**

```bash
# Install dependencies
pip install -r requirements.txt
```

### **Deployment Options:**

#### **Option 1: Internal Web Server (Recommended)**
Point your internal web server to the `dashboard/` folder:

**IIS (Windows Server):**
1. Open IIS Manager
2. Add new site pointing to `dashboard/` folder
3. Set internal URL (e.g., `http://naresco-server/financial-dashboard`)

**Apache/nginx:**
Configure document root to `dashboard/` directory

**Python HTTP Server (Testing):**
```bash
cd dashboard
python -m http.server 8000
# Access at http://localhost:8000
```

#### **Option 2: Vercel (Cloud - Alternative)**
```bash
# One-time setup
git clone https://github.com/charmenlondon-cmd/Naresco-Financial-Reporting.git
cd Naresco-Financial-Reporting
pip install -r requirements.txt

# Deploy
git add dashboard/index.html
git commit -m "Update dashboard"
git push
# Vercel auto-deploys in ~30 seconds
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

# Step 5: Deploy (choose one)
# Internal hosting: Dashboard files auto-update, just refresh browser
# Vercel (optional): git add dashboard/index.html && git commit -m "Update" && git push
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

**"Dashboard not updating"**
- Refresh browser (Ctrl+F5 for hard refresh)
- Check that web server is pointing to correct `dashboard/` folder
- Verify `update-dashboard.bat` completed successfully

**"Admin settings disappeared"**
- Settings are stored per browser, per URL origin
- Different browsers = different settings (this is normal)
- Each user can customize their own view
- To transfer settings: Use browser console to copy `localStorage.getItem('formattingRules')` and paste on new browser with `localStorage.setItem('formattingRules', 'PASTE_HERE')`

**"Git push failed" (if using Vercel)**
- Authenticate with GitHub
- Run `git push` manually to set up credentials

---

## 📞 Support

For issues or questions, contact the Naresco Finance team.

---

**Built for Naresco Finance | Powered by dual-path validation**
