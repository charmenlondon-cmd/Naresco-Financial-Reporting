"""
Budget vs Actual Data Extraction Script

Extracts financial metrics from Excel files and outputs structured JSON data.
Implements name-based metric searching and automatic month detection.
"""

import pandas as pd
import openpyxl
import json
import argparse
from datetime import datetime
from pathlib import Path


# Metric names to search for (from source file)
METRIC_NAMES = [
    "Total Revenue",
    "Total Variable Cost",
    "Contribution Margin",
    "Total Staff Cost (Direct)",
    "Gross Profit / (Loss)",
    "Total Fixed Cost (Indirect)",
    "Net Surplus / (Deflect)"
]

# Abbreviated names for percentage metrics (stored in database)
PERCENTAGE_ABBREVIATIONS = {
    "Total Variable Cost": "TVC %",
    "Contribution Margin": "CM %",
    "Total Staff Cost (Direct)": "TSCD %",
    "Gross Profit / (Loss)": "GP %",
    "Total Fixed Cost (Indirect)": "TFCI %",
    "Net Surplus / (Deflect)": "NSD %"
}

# Dashboard display names
DASHBOARD_NAMES = {
    "Total Revenue": "Total Revnue",
    "Total Variable Cost": "Variable Cost",
    "TVC %": "Variable Cost %",
    "Contribution Margin": "Contribution Margin",
    "CM %": "Contribution Margin %",
    "Total Staff Cost (Direct)": "Fixed Costs (Direct)",
    "TSCD %": "Fixed Costs (Direct) %",
    "Gross Profit / (Loss)": "Gross Profit / Loss",
    "GP %": "Gross Profit / Loss %",
    "Total Fixed Cost (Indirect)": "Fixed Costs (Indirect)",
    "TFCI %": "Fixed Costs (Indirect) %",
    "Net Surplus / (Deflect)": "Net Profit / Loss",
    "NSD %": "Net Profit / Loss %"
}

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]


def detect_month_from_columns(ws, budget_header_row=11):
    """
    Detect which month we're processing by finding "Budget" headers in row 11.

    Each month has Budget/Actual/Variance columns with "Budget" in the header row.

    Args:
        ws: openpyxl worksheet
        budget_header_row: Row containing "Budget" headers (default 11)

    Returns:
        tuple: (month_number, month_name)
    """
    # Find all columns with "Budget" header
    month_count = 0
    max_col = ws.max_column

    for col in range(2, max_col + 1):  # Start from column 2 (after labels)
        cell_value = ws.cell(budget_header_row, col).value
        if cell_value and str(cell_value).strip().lower() == "budget":
            month_count += 1

    if month_count < 1:
        raise ValueError("No 'Budget' headers found in file. Check file structure.")

    month_number = month_count
    month_name = MONTH_NAMES[month_number - 1]

    print(f"Found {month_count} month(s) of data")

    return month_number, month_name


def find_metric_row(ws, metric_name, search_col=1, max_row=150):
    """
    Find the row number where a metric name appears.

    Args:
        ws: openpyxl worksheet
        metric_name: Name of metric to find
        search_col: Column to search in (default 1 = column A)
        max_row: Maximum row to search

    Returns:
        int: Row number, or None if not found
    """
    for row in range(1, max_row + 1):
        cell_value = ws.cell(row, search_col).value
        if cell_value and str(cell_value).strip() == metric_name:
            return row
    return None


def extract_metric_data(ws, metric_name, row_num, month_number):
    """
    Extract Budget, Actual, Variance for a specific metric and month.

    Each month has 3 columns: Budget, Actual, Variance
    Column A = labels (offset 1)
    Month 1 (Jan) = columns 2, 3, 4 (Budget, Actual, Variance)
    Month 2 (Feb) = columns 5, 6, 7
    Month N = columns (N*3 - 1), (N*3), (N*3 + 1)

    Args:
        ws: openpyxl worksheet
        metric_name: Name of the metric
        row_num: Row number where metric is located
        month_number: Which month to extract (1-12)

    Returns:
        dict: {budget, actual, variance}
    """
    # Calculate column positions for this month
    budget_col = (month_number * 3) - 1
    actual_col = month_number * 3
    variance_col = (month_number * 3) + 1

    budget = ws.cell(row_num, budget_col).value
    actual = ws.cell(row_num, actual_col).value
    variance = ws.cell(row_num, variance_col).value

    # Convert to float, handle None values and Excel errors
    def safe_float(value):
        if value is None:
            return 0.0
        if isinstance(value, str) and ('#' in value or value.strip() == ''):
            return 0.0
        try:
            return float(value)
        except (ValueError, TypeError):
            return 0.0

    budget = safe_float(budget)
    actual = safe_float(actual)
    variance = safe_float(variance)

    return {
        "budget": budget,
        "actual": actual,
        "variance": variance
    }


def extract_all_metrics(file_path, sheet_name="Detail Budget"):
    """
    Extract all metrics from the Excel file.

    Args:
        file_path: Path to Excel file
        sheet_name: Name of sheet to extract from

    Returns:
        list: List of metric dictionaries in RAW DATA format
    """
    print(f"Loading workbook: {file_path}")
    wb = openpyxl.load_workbook(file_path, data_only=True)

    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}")

    ws = wb[sheet_name]

    # Detect which month we're processing
    month_number, month_name = detect_month_from_columns(ws)
    print(f"Detected month: {month_name} (Month {month_number})")

    extracted_data = []

    # Extract each metric
    for metric_name in METRIC_NAMES:
        print(f"Searching for metric: {metric_name}")

        # Find the metric row
        row_num = find_metric_row(ws, metric_name)

        if row_num is None:
            print(f"  WARNING: Metric '{metric_name}' not found. Skipping.")
            continue

        print(f"  [OK] Found at row {row_num}")

        # Extract the metric data
        data = extract_metric_data(ws, metric_name, row_num, month_number)

        # Add to results
        extracted_data.append({
            "Month": month_name,
            "DataPoint": metric_name,
            "DashboardName": DASHBOARD_NAMES.get(metric_name, metric_name),
            "Budget": data["budget"],
            "Actual": data["actual"],
            "Variance": data["variance"]
        })

        # Check for percentage in row below
        if metric_name in PERCENTAGE_ABBREVIATIONS:
            percentage_row = row_num + 1
            percentage_cell = ws.cell(percentage_row, 1).value

            # Check if next row contains "%"
            if percentage_cell and "%" in str(percentage_cell):
                print(f"  [OK] Found percentage at row {percentage_row}")

                # Extract percentage data
                pct_data = extract_metric_data(ws, f"{metric_name} %", percentage_row, month_number)
                pct_abbrev = PERCENTAGE_ABBREVIATIONS[metric_name]

                extracted_data.append({
                    "Month": month_name,
                    "DataPoint": pct_abbrev,
                    "DashboardName": DASHBOARD_NAMES.get(pct_abbrev, pct_abbrev),
                    "Budget": pct_data["budget"],
                    "Actual": pct_data["actual"],
                    "Variance": pct_data["variance"]
                })

    wb.close()

    print(f"\n[OK] Extracted {len(extracted_data)} metrics for {month_name}")

    return extracted_data


def save_to_json(data, output_path):
    """
    Save extracted data to JSON file.

    Args:
        data: List of metric dictionaries
        output_path: Path to output JSON file
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w') as f:
        json.dump(data, f, indent=2)

    print(f"\n[OK] Data saved to: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Extract Budget vs Actual data from Excel files"
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Path to input Excel file"
    )
    parser.add_argument(
        "--output",
        default="dashboard/data/raw-data.json",
        help="Path to output JSON file (default: dashboard/data/raw-data.json)"
    )
    parser.add_argument(
        "--sheet",
        default="Detail Budget",
        help="Sheet name to extract from (default: Detail Budget)"
    )

    args = parser.parse_args()

    print("="*60)
    print("BUDGET VS ACTUAL DATA EXTRACTION")
    print("="*60)

    try:
        # Extract data
        data = extract_all_metrics(args.input, args.sheet)

        # Save to JSON
        save_to_json(data, args.output)

        print("\n" + "="*60)
        print("[SUCCESS] EXTRACTION COMPLETE")
        print("="*60)

    except Exception as e:
        print(f"\n[ERROR] {e}")
        raise


if __name__ == "__main__":
    main()
