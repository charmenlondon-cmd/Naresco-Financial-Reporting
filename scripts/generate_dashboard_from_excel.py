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

MONTH_ORDER = ["January","February","March","April","May","June",
               "July","August","September","October","November","December"]


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
        'months': sorted(list(months), key=lambda m: MONTH_ORDER.index(m) if m in MONTH_ORDER else 99),
        'monthly_data': monthly_data,
        'ytd_data': ytd_data,
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }

    print(f"  Loaded {len(months)} months, {len(ytd_data)} YTD metrics")
    return dashboard_data


def export_revenue_json(excel_path, output_path):
    """
    Export Revenue sheet from company database to JSON for dashboard drill-down.
    Output: {months: [...], rows: [{code, particulars, 'Jan-26': val, ..., 'YTD Total': val}]}
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    if 'Revenue' not in wb.sheetnames:
        print("  No Revenue sheet found — skipping revenue detail export")
        wb.close()
        return

    ws = wb['Revenue']

    # Read header row to find months and column count
    headers = []
    for col in range(1, ws.max_column + 1):
        v = ws.cell(1, col).value
        if v is None:
            break
        headers.append(v)

    # headers: CODE | PARTICULARS | Jan-26 | Feb-26 | ... | YTD Total
    month_labels = headers[2:-1]  # everything between PARTICULARS and YTD Total

    rows = []
    for row in range(2, ws.max_row + 1):
        code = ws.cell(row, 1).value
        particulars = ws.cell(row, 2).value
        if code is None and particulars is None:
            continue

        row_data = {
            "code": int(code) if isinstance(code, (int, float)) else None,
            "particulars": str(particulars).strip() if particulars else ""
        }
        for i, month in enumerate(month_labels):
            v = ws.cell(row, 3 + i).value
            row_data[month] = round(float(v), 2) if isinstance(v, (int, float)) else 0.0

        ytd = ws.cell(row, 3 + len(month_labels)).value
        row_data["YTD Total"] = round(float(ytd), 2) if isinstance(ytd, (int, float)) else 0.0

        rows.append(row_data)

    wb.close()

    result = {"months": month_labels, "rows": rows}
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, indent=2)
    print(f"  Revenue detail exported to: {output_path}")


SECTION_SHEETS = {
    "Variable Costs": {"slug": "variable-costs", "exclude": set()},
    "Fixed Costs":    {"slug": "fixed-costs",    "exclude": set()},
    "Indirect Costs": {"slug": "indirect-costs", "exclude": {"Interest Expenses", "Amortization", "Depreciation"}},
    "TDI":            {"slug": "tdi",            "exclude": set()},
}


def export_section_jsons(excel_path, company):
    """
    Export each section sheet from the database to a JSON file for dashboard drill-down.
    Output format:
      {months, rows: [{description, 'Jan-26': {budget, actual, variance}, ..., YTD: {budget, actual, variance}}]}
    """
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception as e:
        print(f"  Could not open database for section export: {e}")
        return

    for sheet_name, cfg in SECTION_SHEETS.items():
        slug = cfg["slug"]
        exclusions = cfg.get("exclude", set())

        if sheet_name not in wb.sheetnames:
            print(f"  No '{sheet_name}' sheet found — skipping")
            continue

        ws = wb[sheet_name]
        headers = []
        for col in range(1, ws.max_column + 1):
            v = ws.cell(1, col).value
            if v is None:
                break
            headers.append(v)

        # Detect month labels
        months = []
        seen = set()
        for h in headers:
            if h and h.endswith(" Budget") and not h.startswith("YTD"):
                label = h[:-7]
                if label not in seen:
                    seen.add(label)
                    months.append(label)

        rows = []
        for row in range(2, ws.max_row + 1):
            desc = ws.cell(row, 1).value
            if not desc:
                continue
            desc_str = str(desc).strip()
            # Skip rows excluded from this section's JSON export
            if desc_str in exclusions:
                continue
            row_data = {"description": desc_str}
            col_offset = 2
            for label in months:
                row_data[label] = {
                    "budget":   _safe(ws.cell(row, col_offset).value),
                    "actual":   _safe(ws.cell(row, col_offset + 1).value),
                    "variance": _safe(ws.cell(row, col_offset + 2).value),
                }
                col_offset += 3
            # YTD columns
            row_data["YTD"] = {
                "budget":   _safe(ws.cell(row, col_offset).value),
                "actual":   _safe(ws.cell(row, col_offset + 1).value),
                "variance": _safe(ws.cell(row, col_offset + 2).value),
            }
            rows.append(row_data)

        result = {"months": months, "rows": rows}
        out = Path(f"dashboard/data/{slug}-{company}.json")
        out.parent.mkdir(parents=True, exist_ok=True)
        with open(out, 'w', encoding='utf-8') as f:
            json.dump(result, f, indent=2)
        print(f"  {sheet_name} detail exported to: {out}")

    wb.close()


def _safe(v):
    return round(float(v), 2) if isinstance(v, (int, float)) else 0.0


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

        # Export detail JSON files
        export_revenue_json(args.excel, f"dashboard/data/revenue-{args.company}.json")
        export_section_jsons(args.excel, args.company)

        print()
        print("=" * 60)
        print("[SUCCESS] DASHBOARD GENERATED FROM EXCEL")
        print("=" * 60)

    except Exception as e:
        print(f"\n[ERROR] {e}")
        raise


if __name__ == "__main__":
    main()
