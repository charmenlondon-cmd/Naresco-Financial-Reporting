"""
Budget vs Actual Data Consolidation Script

Aggregates raw monthly data into YTD (Year-to-Date) totals.
Implements SUMIF logic for absolute values and AVERAGEIF for percentages.
"""

import json
import argparse
from pathlib import Path
from collections import defaultdict


# Percentage metrics use AVERAGE, others use SUM
PERCENTAGE_METRICS = [
    "TVC %",
    "CM %",
    "TSCD %",
    "GP %",
    "TFCI %",
    "NSD %"
]


def load_raw_data(input_path):
    """
    Load raw data from JSON file.

    Args:
        input_path: Path to raw-data.json file

    Returns:
        list: Raw data records
    """
    with open(input_path, 'r') as f:
        data = json.load(f)

    print(f"Loaded {len(data)} records from {input_path}")

    return data


def consolidate_metrics(raw_data):
    """
    Consolidate raw monthly data into YTD totals.

    For absolute values (Revenue, Costs, etc.): SUM across months
    For percentages (%, etc.): AVERAGE across months

    Args:
        raw_data: List of raw data records

    Returns:
        list: Consolidated YTD data
    """
    # Group data by DataPoint
    grouped = defaultdict(lambda: {
        'dashboard_name': '',
        'budget_values': [],
        'actual_values': [],
        'variance_values': []
    })

    for record in raw_data:
        data_point = record['DataPoint']
        grouped[data_point]['dashboard_name'] = record['DashboardName']
        grouped[data_point]['budget_values'].append(record['Budget'])
        grouped[data_point]['actual_values'].append(record['Actual'])
        grouped[data_point]['variance_values'].append(record['Variance'])

    # Calculate YTD totals
    consolidated = []

    for data_point, values in grouped.items():
        # Determine aggregation method
        is_percentage = data_point in PERCENTAGE_METRICS

        if is_percentage:
            # Use AVERAGE for percentages
            ytd_budget = sum(values['budget_values']) / len(values['budget_values']) if values['budget_values'] else 0
            ytd_actual = sum(values['actual_values']) / len(values['actual_values']) if values['actual_values'] else 0
            ytd_variance = ytd_actual - ytd_budget  # Variance = Actual - Budget
            agg_method = "AVERAGE"
        else:
            # Use SUM for absolute values
            ytd_budget = sum(values['budget_values'])
            ytd_actual = sum(values['actual_values'])
            ytd_variance = ytd_actual - ytd_budget  # Variance = Actual - Budget
            agg_method = "SUM"

        consolidated.append({
            'DataPoint': data_point,
            'DashboardName': values['dashboard_name'],
            'YTD_BUDGET': ytd_budget,
            'YTD_ACTUAL': ytd_actual,
            'YTD_VARIANCE': ytd_variance,
            'AggregationMethod': agg_method
        })

        print(f"  [{agg_method}] {data_point}: Budget={ytd_budget:.2f}, Actual={ytd_actual:.2f}, Variance={ytd_variance:.2f}")

    return consolidated


def save_consolidated_data(data, output_path):
    """
    Save consolidated data to JSON file.

    Args:
        data: Consolidated data
        output_path: Path to output JSON file
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, 'w') as f:
        json.dump(data, f, indent=2)

    print(f"\n[OK] Consolidated data saved to: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Consolidate Budget vs Actual raw data into YTD totals"
    )
    parser.add_argument(
        "--input",
        default="dashboard/data/raw-data.json",
        help="Path to raw data JSON file (default: dashboard/data/raw-data.json)"
    )
    parser.add_argument(
        "--output",
        default="dashboard/data/calculations.json",
        help="Path to output consolidated JSON file (default: dashboard/data/calculations.json)"
    )

    args = parser.parse_args()

    print("="*60)
    print("BUDGET VS ACTUAL DATA CONSOLIDATION")
    print("="*60)

    try:
        # Load raw data
        raw_data = load_raw_data(args.input)

        # Get unique months
        months = sorted(set(record['Month'] for record in raw_data))
        print(f"Consolidating data for {len(months)} month(s): {', '.join(months)}")

        # Consolidate metrics
        print("\nCalculating YTD totals:")
        consolidated = consolidate_metrics(raw_data)

        # Save consolidated data
        save_consolidated_data(consolidated, args.output)

        print("\n" + "="*60)
        print("[SUCCESS] CONSOLIDATION COMPLETE")
        print("="*60)
        print(f"Aggregated {len(consolidated)} metrics")

    except Exception as e:
        print(f"\n[ERROR] {e}")
        raise


if __name__ == "__main__":
    main()
