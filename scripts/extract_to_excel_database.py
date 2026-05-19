"""
Budget vs Actual - Extract to Excel Database

Extracts data from source Excel files and appends to Excel database.
Database structure matches Sample Dashboard format.
"""

import pandas as pd
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import argparse
from pathlib import Path
from datetime import datetime


# Metric configurations (from sample dashboard)
METRICS_CONFIG = {
    "Total Revenue": {
        "dashboard_name": "Total Revnue",
        "row_name": "Total Revenue",
        "has_percentage": False
    },
    "Total Variable Cost": {
        "dashboard_name": "Variable Cost",
        "row_name": "Total Variable Cost",
        "has_percentage": True,
        "percentage_abbrev": "TVC %",
        "percentage_dashboard": "Variable Cost %"
    },
    "Contribution Margin": {
        "dashboard_name": "Contribution Margin",
        "row_name": "Contribution Margin",
        "has_percentage": True,
        "percentage_abbrev": "CM %",
        "percentage_dashboard": "Contribution Margin %"
    },
    "Total Staff Cost (Direct)": {
        "dashboard_name": "Fixed Costs (Direct)",
        "row_name": "Total Staff Cost (Direct)",
        "has_percentage": True,
        "percentage_abbrev": "TSCD %",
        "percentage_dashboard": "Fixed Costs (Direct) %"
    },
    "Gross Profit / (Loss)": {
        "dashboard_name": "Gross Profit / Loss",
        "row_name": "Gross Profit / (Loss)",
        "has_percentage": True,
        "percentage_abbrev": "GP %",
        "percentage_dashboard": "Gross Profit / Loss %"
    },
    "Total Fixed Cost (Indirect)": {
        "dashboard_name": "Fixed Costs (Indirect)",
        "row_name": "Total Fixed Cost (Indirect)",
        "has_percentage": True,
        "percentage_abbrev": "TFCI %",
        "percentage_dashboard": "Fixed Costs (Indirect) %"
    },
    "Net Surplus / (Deflect)": {
        "dashboard_name": "Net Profit / Loss",
        "row_name": "Net Surplus / (Deflect)",
        "has_percentage": True,
        "percentage_abbrev": "NSD %",
        "percentage_dashboard": "Net Profit / Loss %"
    },
    "Interest Expenses": {
        "dashboard_name": "Interest Expenses",
        "row_name": "Interest Expenses",
        "has_percentage": False
    },
    "Amortization": {
        "dashboard_name": "Amortization",
        "row_name": "Amortization",
        "has_percentage": False
    },
    "Depreciation": {
        "dashboard_name": "Depreciation",
        "row_name": "Depreciation",
        "has_percentage": False
    },
}

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


def detect_column_structure(ws):
    """
    Dynamically detect Budget/Actual/Variance column positions for each month.
    Scans rows 1-25 to find the Budget header row, then maps column groups.
    Returns list of (month_num, budget_col, actual_col, variance_col).
    """
    header_row = None
    for row in range(1, 26):
        for col in range(2, min(ws.max_column + 1, 15)):
            v = ws.cell(row, col).value
            if v and str(v).strip().lower() == "budget":
                header_row = row
                break
        if header_row:
            break

    if not header_row:
        raise ValueError("No 'Budget' header row found in first 25 rows. Check file structure.")

    months = []
    month_num = 0

    for col in range(2, ws.max_column + 1):
        v = ws.cell(header_row, col).value
        if not v or str(v).strip().lower() != "budget":
            continue

        budget_col = col
        actual_col = None
        variance_col = None

        for ahead in range(col + 1, min(col + 7, ws.max_column + 2)):
            av = ws.cell(header_row, ahead).value
            if av is None:
                continue
            av_str = str(av).strip().lower()
            if "actual" in av_str and actual_col is None:
                actual_col = ahead
            elif "variance" in av_str and actual_col is not None:
                variance_col = ahead
                break
            elif av_str == "budget":
                break

        if actual_col and variance_col:
            month_num += 1
            months.append((month_num, budget_col, actual_col, variance_col))

    if not months:
        raise ValueError("No complete Budget/Actual/Variance column groups found.")

    return months


def detect_months_in_file(ws, budget_header_row=11):
    """Detect how many months of data are in the file using dynamic detection."""
    col_structure = detect_column_structure(ws)
    return len(col_structure)


def find_metric_row(ws, metric_name, search_col=1, max_row=150):
    """Find row number for a metric name."""
    for row in range(1, max_row + 1):
        cell_value = ws.cell(row, search_col).value
        if cell_value:
            cell_str = str(cell_value).strip()
            if cell_str == metric_name or cell_str.startswith(metric_name):
                return row
    return None


def extract_metric_value(ws, row_num, col):
    """Extract a single value from the worksheet at (row_num, col)."""
    value = ws.cell(row_num, col).value
    if value is None:
        return 0.0
    if isinstance(value, str) and '#' in value:
        return 0.0
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0.0


def extract_from_source(file_path, sheet_name="Detail Budget"):
    """
    Extract all data from source Excel file.
    Returns DataFrame ready to append to BVA_DATA sheet.
    """
    print(f"Loading source file: {file_path}")
    wb = openpyxl.load_workbook(file_path, data_only=True)

    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found")

    ws = wb[sheet_name]

    # Detect months and column positions
    col_structure = detect_column_structure(ws)
    total_months = len(col_structure)
    print(f"Detected {total_months} month(s) of data")

    # Extract data for all months
    records = []

    for month_num in range(1, total_months + 1):
        month_name = MONTH_NAMES[month_num - 1]
        print(f"  Extracting {month_name}...")

        for metric_key, config in METRICS_CONFIG.items():
            row_num = find_metric_row(ws, config["row_name"])

            if row_num is None:
                print(f"    WARNING: '{config['row_name']}' not found")
                continue

            # Extract Budget, Actual, Variance
            month_entry = next((m for m in col_structure if m[0] == month_num), None)
            if not month_entry:
                continue
            _, budget_col, actual_col, variance_col = month_entry
            budget = extract_metric_value(ws, row_num, budget_col)
            actual = extract_metric_value(ws, row_num, actual_col)
            variance = extract_metric_value(ws, row_num, variance_col)

            # Add main metric
            records.append({
                'Month': month_name,
                'DataPoint': metric_key,
                'Sample Data ROW Number': row_num,
                'DashboardName': config['dashboard_name'],
                'Budget': budget,
                'Actual': actual,
                'Variance': variance
            })

            # Add percentage if applicable
            if config.get('has_percentage'):
                pct_row = row_num + 1
                pct_budget = extract_metric_value(ws, pct_row, budget_col)
                pct_actual = extract_metric_value(ws, pct_row, actual_col)
                pct_variance = extract_metric_value(ws, pct_row, variance_col)

                records.append({
                    'Month': month_name,
                    'DataPoint': config['percentage_abbrev'],
                    'Sample Data ROW Number': pct_row,
                    'DashboardName': config['percentage_dashboard'],
                    'Budget': pct_budget,
                    'Actual': pct_actual,
                    'Variance': pct_variance
                })

    wb.close()

    df = pd.DataFrame(records)
    print(f"\n[OK] Extracted {len(records)} records ({total_months} months x 13 metrics)")

    return df


def append_to_database(df, database_path):
    """
    Append extracted data to Excel database BVA_DATA sheet.
    Creates database if it doesn't exist.
    """
    database_path = Path(database_path)

    # Check if database exists
    if database_path.exists():
        print(f"\nAppending to existing database: {database_path}")
        wb = load_workbook(database_path)

        # Get or create BVA_DATA sheet
        if 'BVA_DATA' in wb.sheetnames:
            ws = wb['BVA_DATA']
            start_row = ws.max_row + 1
        else:
            ws = wb.create_sheet('BVA_DATA', 0)
            start_row = 1
    else:
        print(f"\nCreating new database: {database_path}")
        wb = Workbook()
        ws = wb.active
        ws.title = 'BVA_DATA'
        start_row = 1

    # Write data
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=(start_row == 1)), start_row):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # Save
    wb.save(database_path)
    wb.close()

    print(f"[OK] Data appended to BVA_DATA sheet (starting row {start_row})")
    print(f"[OK] Database saved: {database_path}")


def create_calculations_sheet(database_path):
    """
    Create BVA_CALC sheet with formulas (SUMIF/AVERAGEIF).
    """
    wb = load_workbook(database_path)

    # Create BVA_CALC sheet if it doesn't exist
    if 'BVA_CALC' in wb.sheetnames:
        ws = wb['BVA_CALC']
        ws.delete_rows(1, ws.max_row)  # Clear existing
    else:
        ws = wb.create_sheet('BVA_CALC')

    # Headers
    ws.append(['DataPoint', 'DashboardName', 'YTD_BUDGET', 'YTD_ACTUAL', 'YTD_VARIANCE'])

    # Metrics (order matters for dashboard)
    metrics_order = list(METRICS_CONFIG.keys())
    percentage_metrics = ['TVC %', 'CM %', 'TSCD %', 'GP %', 'TFCI %', 'NSD %']

    row = 2
    for metric in metrics_order:
        config = METRICS_CONFIG[metric]

        # Main metric (SUM)
        ws.cell(row, 1, metric)
        ws.cell(row, 2, config['dashboard_name'])
        ws.cell(row, 3, f'=SUMIF(\'BVA_DATA\'!B:B,A{row},\'BVA_DATA\'!E:E)')
        ws.cell(row, 4, f'=SUMIF(\'BVA_DATA\'!B:B,A{row},\'BVA_DATA\'!F:F)')
        ws.cell(row, 5, f'=SUMIF(\'BVA_DATA\'!B:B,A{row},\'BVA_DATA\'!G:G)')
        row += 1

        # Percentage (AVERAGE)
        if config.get('has_percentage'):
            pct_abbrev = config['percentage_abbrev']
            pct_dashboard = config['percentage_dashboard']
            ws.cell(row, 1, pct_abbrev)
            ws.cell(row, 2, pct_dashboard)
            ws.cell(row, 3, f'=AVERAGEIF(\'BVA_DATA\'!B:B,A{row},\'BVA_DATA\'!E:E)')
            ws.cell(row, 4, f'=AVERAGEIF(\'BVA_DATA\'!B:B,A{row},\'BVA_DATA\'!F:F)')
            ws.cell(row, 5, f'=AVERAGEIF(\'BVA_DATA\'!B:B,A{row},\'BVA_DATA\'!G:G)')
            row += 1

    # EBITDA (derived: Net Surplus + Interest + Amortization + Depreciation)
    ebitda_components = [
        "Net Surplus / (Deflect)",
        "Interest Expenses",
        "Amortization",
        "Depreciation"
    ]
    sumif_budget = "+".join([f"SUMIF('BVA_DATA'!B:B,\"{m}\",'BVA_DATA'!E:E)" for m in ebitda_components])
    sumif_actual = "+".join([f"SUMIF('BVA_DATA'!B:B,\"{m}\",'BVA_DATA'!F:F)" for m in ebitda_components])

    ws.cell(row, 1, "EBITDA")
    ws.cell(row, 2, "EBITDA")
    ws.cell(row, 3, f"={sumif_budget}")
    ws.cell(row, 4, f"={sumif_actual}")
    ws.cell(row, 5, f"=D{row}-C{row}")
    ebitda_row = row
    row += 1

    # EBITDA %
    ws.cell(row, 1, "EBITDA %")
    ws.cell(row, 2, "EBITDA %")
    ws.cell(row, 3, f"=IF(SUMIF('BVA_DATA'!B:B,\"Total Revenue\",'BVA_DATA'!E:E)=0,0,C{ebitda_row}/SUMIF('BVA_DATA'!B:B,\"Total Revenue\",'BVA_DATA'!E:E))")
    ws.cell(row, 4, f"=IF(SUMIF('BVA_DATA'!B:B,\"Total Revenue\",'BVA_DATA'!F:F)=0,0,D{ebitda_row}/SUMIF('BVA_DATA'!B:B,\"Total Revenue\",'BVA_DATA'!F:F))")
    ws.cell(row, 5, f"=D{row}-C{row}")
    row += 1

    wb.save(database_path)
    wb.close()

    print(f"[OK] BVA_CALC sheet created with formulas")


def main():
    parser = argparse.ArgumentParser(
        description="Extract Budget vs Actual data to Excel database"
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Path to source Excel file"
    )
    parser.add_argument(
        '--company',
        required=True,
        help='Company ID (e.g. mudin, mantis)'
    )
    parser.add_argument(
        "--database",
        default=None,
        help="Path to Excel database"
    )
    parser.add_argument(
        "--sheet",
        default="Detail Budget",
        help="Source sheet name (default: Detail Budget)"
    )

    args = parser.parse_args()

    if args.database is None:
        args.database = f"Financial-Data-Database-{args.company.capitalize()}.xlsx"

    print("="*60)
    print("EXTRACT TO EXCEL DATABASE")
    print("="*60)

    try:
        # Extract from source
        df = extract_from_source(args.input, args.sheet)

        # Append to database
        append_to_database(df, args.database)

        # Create/update BVA_CALC sheet
        create_calculations_sheet(args.database)

        print("\n" + "="*60)
        print("[SUCCESS] DATA EXTRACTED TO EXCEL DATABASE")
        print("="*60)
        print(f"\nDatabase: {args.database}")
        print("  - BVA_DATA sheet: Accumulated monthly data")
        print("  - BVA_CALC sheet: YTD formulas (SUMIF/AVERAGEIF)")

    except Exception as e:
        print(f"\n[ERROR] {e}")
        raise


if __name__ == "__main__":
    main()
