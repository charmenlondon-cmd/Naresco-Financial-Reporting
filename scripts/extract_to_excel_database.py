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
    }
}

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


def detect_months_in_file(ws, budget_header_row=11):
    """
    Detect how many months of data are in the file.
    Counts 'Budget' headers in row 11.
    """
    month_count = 0
    for col in range(2, ws.max_column + 1):
        cell_value = ws.cell(budget_header_row, col).value
        if cell_value and str(cell_value).strip().lower() == "budget":
            month_count += 1
    return month_count


def find_metric_row(ws, metric_name, search_col=1, max_row=150):
    """Find row number for a metric name."""
    for row in range(1, max_row + 1):
        cell_value = ws.cell(row, search_col).value
        if cell_value and str(cell_value).strip() == metric_name:
            return row
    return None


def extract_metric_value(ws, row_num, month_num, value_type='budget'):
    """
    Extract Budget, Actual, or Variance for a metric and month.
    month_num: 1=Jan, 2=Feb, etc.
    value_type: 'budget', 'actual', or 'variance'
    """
    # Column calculation: each month has 3 columns (Budget, Actual, Variance)
    # Month 1: cols 2, 3, 4
    # Month 2: cols 5, 6, 7
    # Month N: cols (N*3-1), (N*3), (N*3+1)

    if value_type == 'budget':
        col = (month_num * 3) - 1
    elif value_type == 'actual':
        col = month_num * 3
    elif value_type == 'variance':
        col = (month_num * 3) + 1
    else:
        raise ValueError(f"Invalid value_type: {value_type}")

    value = ws.cell(row_num, col).value

    # Handle errors and None
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
    Returns DataFrame ready to append to RAW DATA sheet.
    """
    print(f"Loading source file: {file_path}")
    wb = openpyxl.load_workbook(file_path, data_only=True)

    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found")

    ws = wb[sheet_name]

    # Detect months
    total_months = detect_months_in_file(ws)
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
            budget = extract_metric_value(ws, row_num, month_num, 'budget')
            actual = extract_metric_value(ws, row_num, month_num, 'actual')
            variance = extract_metric_value(ws, row_num, month_num, 'variance')

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
                pct_budget = extract_metric_value(ws, pct_row, month_num, 'budget')
                pct_actual = extract_metric_value(ws, pct_row, month_num, 'actual')
                pct_variance = extract_metric_value(ws, pct_row, month_num, 'variance')

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
    Append extracted data to Excel database RAW DATA sheet.
    Creates database if it doesn't exist.
    """
    database_path = Path(database_path)

    # Check if database exists
    if database_path.exists():
        print(f"\nAppending to existing database: {database_path}")
        wb = load_workbook(database_path)

        # Get or create RAW DATA sheet
        if 'RAW DATA' in wb.sheetnames:
            ws = wb['RAW DATA']
            start_row = ws.max_row + 1
        else:
            ws = wb.create_sheet('RAW DATA', 0)
            start_row = 1
    else:
        print(f"\nCreating new database: {database_path}")
        wb = Workbook()
        ws = wb.active
        ws.title = 'RAW DATA'
        start_row = 1

    # Write data
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=(start_row == 1)), start_row):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # Save
    wb.save(database_path)
    wb.close()

    print(f"[OK] Data appended to RAW DATA sheet (starting row {start_row})")
    print(f"[OK] Database saved: {database_path}")


def create_calculations_sheet(database_path):
    """
    Create CALCULATIONS sheet with formulas (SUMIF/AVERAGEIF).
    """
    wb = load_workbook(database_path)

    # Create CALCULATIONS sheet if it doesn't exist
    if 'CALCULATIONS' in wb.sheetnames:
        ws = wb['CALCULATIONS']
        ws.delete_rows(1, ws.max_row)  # Clear existing
    else:
        ws = wb.create_sheet('CALCULATIONS')

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
        ws.cell(row, 3, f'=SUMIF(\'RAW DATA\'!B:B,A{row},\'RAW DATA\'!E:E)')
        ws.cell(row, 4, f'=SUMIF(\'RAW DATA\'!B:B,A{row},\'RAW DATA\'!F:F)')
        ws.cell(row, 5, f'=SUMIF(\'RAW DATA\'!B:B,A{row},\'RAW DATA\'!G:G)')
        row += 1

        # Percentage (AVERAGE)
        if config.get('has_percentage'):
            pct_abbrev = config['percentage_abbrev']
            pct_dashboard = config['percentage_dashboard']
            ws.cell(row, 1, pct_abbrev)
            ws.cell(row, 2, pct_dashboard)
            ws.cell(row, 3, f'=AVERAGEIF(\'RAW DATA\'!B:B,A{row},\'RAW DATA\'!E:E)')
            ws.cell(row, 4, f'=AVERAGEIF(\'RAW DATA\'!B:B,A{row},\'RAW DATA\'!F:F)')
            ws.cell(row, 5, f'=AVERAGEIF(\'RAW DATA\'!B:B,A{row},\'RAW DATA\'!G:G)')
            row += 1

    wb.save(database_path)
    wb.close()

    print(f"[OK] CALCULATIONS sheet created with formulas")


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
        "--database",
        default="Financial-Data-Database.xlsx",
        help="Path to Excel database (default: Financial-Data-Database.xlsx)"
    )
    parser.add_argument(
        "--sheet",
        default="Detail Budget",
        help="Source sheet name (default: Detail Budget)"
    )

    args = parser.parse_args()

    print("="*60)
    print("EXTRACT TO EXCEL DATABASE")
    print("="*60)

    try:
        # Extract from source
        df = extract_from_source(args.input, args.sheet)

        # Append to database
        append_to_database(df, args.database)

        # Create/update CALCULATIONS sheet
        create_calculations_sheet(args.database)

        print("\n" + "="*60)
        print("[SUCCESS] DATA EXTRACTED TO EXCEL DATABASE")
        print("="*60)
        print(f"\nDatabase: {args.database}")
        print("  - RAW DATA sheet: Accumulated monthly data")
        print("  - CALCULATIONS sheet: YTD formulas (SUMIF/AVERAGEIF)")

    except Exception as e:
        print(f"\n[ERROR] {e}")
        raise


if __name__ == "__main__":
    main()
