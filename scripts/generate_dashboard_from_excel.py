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


def read_excel_database(excel_path, calculations_json_path):
    """
    Read monthly data from BVA_DATA sheet and YTD totals from calculations JSON.
    BVA_DATA columns: Month | DataPoint | SampleRowNum | DashboardName | Budget | Actual | Variance
    """
    print(f"Reading data from {excel_path}...")

    wb = openpyxl.load_workbook(excel_path, data_only=True)

    if 'BVA_DATA' not in wb.sheetnames:
        raise ValueError("BVA_DATA sheet not found in Excel database")

    ws_data = wb['BVA_DATA']
    monthly_data = {}
    months = set()

    for row in range(2, ws_data.max_row + 1):
        month = ws_data.cell(row, 1).value
        data_point = ws_data.cell(row, 2).value
        # col 3 = Sample Data ROW Number (skip)
        dashboard_name = ws_data.cell(row, 4).value
        budget = ws_data.cell(row, 5).value
        actual = ws_data.cell(row, 6).value
        variance = ws_data.cell(row, 7).value

        if month and data_point:
            months.add(month)
            if month not in monthly_data:
                monthly_data[month] = {}
            monthly_data[month][data_point] = {
                'dashboard_name': dashboard_name or data_point,
                'budget': float(budget) if isinstance(budget, (int, float)) else 0.0,
                'actual': float(actual) if isinstance(actual, (int, float)) else 0.0,
                'variance': float(variance) if isinstance(variance, (int, float)) else 0.0
            }

    wb.close()

    # Read YTD from calculations JSON (already validated and computed)
    print(f"Reading YTD totals from {calculations_json_path}...")
    with open(calculations_json_path, 'r') as f:
        calc_data = json.load(f)

    ytd_data = {}
    for item in calc_data:
        metric = item.get('DataPoint')
        if metric:
            ytd_data[metric] = {
                'dashboard_name': item.get('DashboardName', metric),
                'budget': item.get('YTD_BUDGET', 0.0),
                'actual': item.get('YTD_ACTUAL', 0.0),
                'variance': item.get('YTD_VARIANCE', 0.0)
            }

    dashboard_data = {
        'months': sorted(list(months)),
        'monthly_data': monthly_data,
        'ytd_data': ytd_data,
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    print(f"  Loaded {len(months)} months, {len(ytd_data)} YTD metrics")
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
        '--company',
        required=True,
        help='Company ID (e.g. mudin, mantis)'
    )
    parser.add_argument(
        "--excel",
        default=None,
        help="Path to Excel database"
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Path to output JSON file"
    )

    args = parser.parse_args()

    if args.excel is None:
        args.excel = f"Financial-Data-Database-{args.company.capitalize()}.xlsx"
    if args.output is None:
        args.output = f"dashboard/data/dashboard-{args.company}.json"
    calculations_json = f"dashboard/data/calculations-{args.company}.json"

    print("=" * 60)
    print("DASHBOARD GENERATION (Excel-Based)")
    print("=" * 60)
    print()

    try:
        # Read monthly data from Excel, YTD from calculations JSON
        dashboard_data = read_excel_database(args.excel, calculations_json)

        import json
        from pathlib import Path
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(dashboard_data, f, indent=2)
        print(f"  Dashboard data saved to: {output_path}")

        print()
        print("=" * 60)
        print("[SUCCESS] DASHBOARD GENERATED FROM EXCEL")
        print("=" * 60)

    except Exception as e:
        print(f"\n[ERROR] {e}")
        raise


if __name__ == "__main__":
    main()
