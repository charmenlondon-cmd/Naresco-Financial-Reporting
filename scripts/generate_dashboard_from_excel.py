"""
Budget vs Actual Dashboard Generator (Excel-Based)

Generates dashboard HTML from Excel database (primary source).
Reads directly from Financial-Data-Database.xlsx.
"""

import json
import openpyxl
import argparse
from pathlib import Path
from datetime import datetime


def read_excel_database(excel_path="Financial-Data-Database.xlsx"):
    """
    Read data from Excel database (BVA_DATA and BVA_CALC sheets).

    Returns:
        dict: Dashboard data structure
    """
    print(f"Reading data from {excel_path}...")

    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # Read raw monthly data from BVA_DATA
    if 'BVA_DATA' not in wb.sheetnames:
        raise ValueError("BVA_DATA sheet not found in Excel database")

    ws_data = wb['BVA_DATA']

    monthly_data = {}
    months = set()

    # Read header row (row 1)
    # Expected: Month | DataPoint | DashboardName | Budget | Actual | Variance

    for row in range(2, ws_data.max_row + 1):
        month = ws_data.cell(row, 1).value
        data_point = ws_data.cell(row, 2).value
        dashboard_name = ws_data.cell(row, 3).value
        budget = ws_data.cell(row, 4).value
        actual = ws_data.cell(row, 5).value
        variance = ws_data.cell(row, 6).value

        if month and data_point:
            months.add(month)

            if month not in monthly_data:
                monthly_data[month] = {}

            monthly_data[month][data_point] = {
                'dashboard_name': dashboard_name or data_point,
                'budget': float(budget) if budget else 0.0,
                'actual': float(actual) if actual else 0.0,
                'variance': float(variance) if variance else 0.0
            }

    # Read YTD calculations from BVA_CALC
    if 'BVA_CALC' not in wb.sheetnames:
        raise ValueError("BVA_CALC sheet not found in Excel database")

    ws_calc = wb['BVA_CALC']

    ytd_data = {}

    # Read calculations (starting from row 2)
    # Expected: Metric | Method | YTD Budget | YTD Actual | YTD Variance

    for row in range(2, ws_calc.max_row + 1):
        metric = ws_calc.cell(row, 1).value
        method = ws_calc.cell(row, 2).value
        ytd_budget = ws_calc.cell(row, 3).value
        ytd_actual = ws_calc.cell(row, 4).value
        ytd_variance = ws_calc.cell(row, 5).value

        if metric:
            # Get dashboard name from first occurrence in monthly data
            dashboard_name = metric
            for month_data in monthly_data.values():
                if metric in month_data:
                    dashboard_name = month_data[metric]['dashboard_name']
                    break

            ytd_data[metric] = {
                'dashboard_name': dashboard_name,
                'budget': float(ytd_budget) if ytd_budget else 0.0,
                'actual': float(ytd_actual) if ytd_actual else 0.0,
                'variance': float(ytd_variance) if ytd_variance else 0.0
            }

    wb.close()

    # Prepare dashboard data structure
    dashboard_data = {
        'months': sorted(list(months)),
        'monthly_data': monthly_data,
        'ytd_data': ytd_data,
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    print(f"  Loaded {len(months)} months")
    print(f"  Loaded {len(ytd_data)} YTD metrics")

    return dashboard_data


def generate_dashboard(dashboard_data, template_path, output_path):
    """
    Generate dashboard HTML from Excel data.

    Args:
        dashboard_data: Data dictionary from Excel
        template_path: Path to HTML template
        output_path: Path to output HTML file
    """
    print("Generating dashboard HTML...")

    # Read template
    template_path = Path(template_path)
    with open(template_path, 'r', encoding='utf-8') as f:
        template = f.read()

    # Inject data into template
    dashboard_json = json.dumps(dashboard_data, indent=2)
    html = template.replace('{{DASHBOARD_DATA}}', dashboard_json)

    # Write output
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"  Dashboard saved to: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Generate Budget vs Actual dashboard from Excel database"
    )
    parser.add_argument(
        "--excel",
        default="Financial-Data-Database.xlsx",
        help="Path to Excel database (default: Financial-Data-Database.xlsx)"
    )
    parser.add_argument(
        "--template",
        default="templates/dashboard_template.html",
        help="Path to dashboard template (default: templates/dashboard_template.html)"
    )
    parser.add_argument(
        "--output",
        default="dashboard/index.html",
        help="Path to output HTML (default: dashboard/index.html)"
    )

    args = parser.parse_args()

    print("=" * 60)
    print("DASHBOARD GENERATION (Excel-Based)")
    print("=" * 60)
    print()

    try:
        # Read from Excel database
        dashboard_data = read_excel_database(args.excel)

        # Generate dashboard
        generate_dashboard(dashboard_data, args.template, args.output)

        print()
        print("=" * 60)
        print("[SUCCESS] DASHBOARD GENERATED FROM EXCEL")
        print("=" * 60)

    except Exception as e:
        print(f"\n[ERROR] {e}")
        raise


if __name__ == "__main__":
    main()
