"""
Budget vs Actual Dashboard Generation Script

Generates interactive HTML dashboard from consolidated data.
Includes executive KPIs, trend charts, and variance analysis.
"""

import json
import argparse
from pathlib import Path
from datetime import datetime


def load_data(raw_data_path, calculations_path):
    """
    Load raw and consolidated data.

    Args:
        raw_data_path: Path to raw-data.json
        calculations_path: Path to calculations.json

    Returns:
        tuple: (raw_data, calculations_data)
    """
    with open(raw_data_path, 'r') as f:
        raw_data = json.load(f)

    with open(calculations_path, 'r') as f:
        calculations = json.load(f)

    print(f"Loaded {len(raw_data)} raw records")
    print(f"Loaded {len(calculations)} calculated metrics")

    return raw_data, calculations


def prepare_dashboard_data(raw_data, calculations):
    """
    Prepare data structures for dashboard rendering.

    Args:
        raw_data: Raw monthly data
        calculations: YTD calculations

    Returns:
        dict: Dashboard data structure
    """
    # Get unique months sorted chronologically
    months = sorted(set(record['Month'] for record in raw_data),
                   key=lambda m: ["January", "February", "March", "April", "May", "June",
                                 "July", "August", "September", "October", "November", "December"].index(m))

    # Organize raw data by month and metric
    monthly_data = {}
    for record in raw_data:
        month = record['Month']
        if month not in monthly_data:
            monthly_data[month] = {}
        monthly_data[month][record['DataPoint']] = {
            'budget': record['Budget'],
            'actual': record['Actual'],
            'variance': record['Variance'],
            'dashboard_name': record['DashboardName']
        }

    # Organize calculations by metric
    ytd_data = {}
    for calc in calculations:
        ytd_data[calc['DataPoint']] = {
            'budget': calc['YTD_BUDGET'],
            'actual': calc['YTD_ACTUAL'],
            'variance': calc['YTD_VARIANCE'],
            'dashboard_name': calc['DashboardName']
        }

    return {
        'months': months,
        'monthly_data': monthly_data,
        'ytd_data': ytd_data,
        'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }


def generate_dashboard_html(dashboard_data, template_path, output_path):
    """
    Generate HTML dashboard from template and data.

    Args:
        dashboard_data: Prepared dashboard data
        template_path: Path to HTML template
        output_path: Path to save generated dashboard
    """
    # Read template
    with open(template_path, 'r', encoding='utf-8') as f:
        template = f.read()

    # Convert data to JSON for JavaScript
    data_json = json.dumps(dashboard_data, indent=2)

    # Replace placeholder in template
    html = template.replace('{{DASHBOARD_DATA}}', data_json)

    # Save generated dashboard
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\n[OK] Dashboard generated: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Generate Budget vs Actual HTML dashboard"
    )
    parser.add_argument(
        "--raw-data",
        default="dashboard/data/raw-data.json",
        help="Path to raw data JSON (default: dashboard/data/raw-data.json)"
    )
    parser.add_argument(
        "--calculations",
        default="dashboard/data/calculations.json",
        help="Path to calculations JSON (default: dashboard/data/calculations.json)"
    )
    parser.add_argument(
        "--template",
        default="templates/dashboard_template.html",
        help="Path to HTML template (default: templates/dashboard_template.html)"
    )
    parser.add_argument(
        "--output",
        default="dashboard/index.html",
        help="Path to output HTML file (default: dashboard/index.html)"
    )

    args = parser.parse_args()

    print("="*60)
    print("DASHBOARD GENERATION")
    print("="*60)

    try:
        # Load data
        raw_data, calculations = load_data(args.raw_data, args.calculations)

        # Prepare dashboard data
        print("\nPreparing dashboard data...")
        dashboard_data = prepare_dashboard_data(raw_data, calculations)
        print(f"  Months: {', '.join(dashboard_data['months'])}")
        print(f"  YTD Metrics: {len(dashboard_data['ytd_data'])}")

        # Generate HTML
        print("\nGenerating HTML dashboard...")
        generate_dashboard_html(dashboard_data, args.template, args.output)

        print("\n" + "="*60)
        print("[SUCCESS] DASHBOARD GENERATION COMPLETE")
        print("="*60)
        print(f"Open {args.output} in your browser to view the dashboard")

    except Exception as e:
        print(f"\n[ERROR] {e}")
        raise


if __name__ == "__main__":
    main()
