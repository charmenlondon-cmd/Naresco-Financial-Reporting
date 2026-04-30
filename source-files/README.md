# Source Files Folder

## How to Update the Dashboard

### Quick Start (One-Click Update)

1. **Drop your new Excel file here** (e.g., "Mudin Budget vs Actual File 2026 April.xlsx")
2. **Double-click `update-dashboard.bat`** (in the parent folder)
3. **Wait ~1 minute** - the script will:
   - Find the newest Excel file in this folder
   - Extract the financial data
   - Calculate YTD totals
   - Generate the updated dashboard
   - Deploy to the web automatically
4. **Done!** Dashboard live at https://naresco-financial-reporting.vercel.app

### Important Notes

- The script always processes the **newest** .xlsx file in this folder
- You can keep old files here - they won't interfere
- File name doesn't matter - script finds it by date
- Expected format: "Detail Budget" sheet with Budget/Actual/Variance columns

### What Happens Behind the Scenes

```
source-files/April-file.xlsx
    ↓
[Extract] → dashboard/data/raw-data.json (all months: Jan, Feb, Mar, Apr)
    ↓
[Consolidate] → dashboard/data/calculations.json (YTD totals)
    ↓
[Generate] → dashboard/index.html (updated dashboard)
    ↓
[Deploy] → Git push → Vercel auto-deploy → Live dashboard
```

### Troubleshooting

**"No Excel files found"**
- Make sure you dropped an .xlsx file in this folder

**"Data extraction failed"**
- Check that the Excel file has a "Detail Budget" sheet
- Verify the file structure matches the expected format

**"Git push failed"**
- You may need to authenticate with GitHub
- Run `git push` manually in the terminal to set up credentials

### Manual Process (If Needed)

If you prefer to run commands manually:

```bash
# 1. Extract data
python scripts/extract_budget_data.py --input "source-files/your-file.xlsx"

# 2. Consolidate
python scripts/consolidate_data.py

# 3. Generate dashboard
python scripts/generate_dashboard.py

# 4. Deploy
git add dashboard/index.html
git commit -m "Update dashboard with April data"
git push
```

---

**Need help?** Contact your friendly neighborhood Claude! 🤖
