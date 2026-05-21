"""
Budget vs Actual Data Consolidation Script

Aggregates raw monthly data into YTD (Year-to-Date) totals.
Absolute metrics: SUM across months.
Percentage metrics: derived from YTD numerator / YTD denominator (not averaged).
"""

import json
import argparse
from pathlib import Path
from collections import defaultdict


# Percentage metrics are derived, not summed/averaged
PERCENTAGE_METRICS = ["TVC %", "CM %", "TSCD %", "GP %", "TFCI %", "NSD %", "EBITDA %"]

# How each % is derived: (numerator metric, denominator metric, variance sign)
# variance sign: 'cost' = Budget-Actual (positive=underspend=good)
#                'profit' = Actual-Budget (positive=above budget=good)
PCT_DERIVATIONS = {
    "TVC %":  ("Total Variable Cost",      "Total Revenue", "cost"),
    "CM %":   ("Contribution Margin",       "Total Revenue", "profit"),
    "TSCD %": ("Total Staff Cost (Direct)", "Total Revenue", "cost"),
    "GP %":   ("Gross Profit / (Loss)",     "Total Revenue", "profit"),
    "TFCI %": ("Adjusted Indirect Costs",   "Total Revenue", "cost"),   # uses adjusted (excl. TDI)
    "NSD %":  ("Net Surplus / (Deflect)",   "Total Revenue", "profit"),
}

TDI_COMPONENTS = ["Interest Expenses", "Amortization", "Depreciation"]


def load_raw_data(input_path):
    with open(input_path, 'r') as f:
        data = json.load(f)
    print(f"Loaded {len(data)} records from {input_path}")
    return data


def consolidate_metrics(raw_data):
    """
    Consolidate raw monthly data into YTD totals.
    Absolute metrics are summed; percentage metrics are derived from YTD absolutes.
    """
    grouped = defaultdict(lambda: {
        'dashboard_name': '',
        'budget_values': [],
        'actual_values': [],
        'variance_values': []
    })

    for record in raw_data:
        dp = record['DataPoint']
        grouped[dp]['dashboard_name'] = record['DashboardName']
        grouped[dp]['budget_values'].append(record['Budget'])
        grouped[dp]['actual_values'].append(record['Actual'])
        grouped[dp]['variance_values'].append(record['Variance'])

    consolidated = []

    # Step 1: SUM all absolute (non-percentage) metrics
    for data_point, values in grouped.items():
        if data_point in PERCENTAGE_METRICS:
            continue  # derived in step 2

        ytd_budget   = sum(values['budget_values'])
        ytd_actual   = sum(values['actual_values'])
        ytd_variance = sum(values['variance_values'])

        consolidated.append({
            'DataPoint':         data_point,
            'DashboardName':     values['dashboard_name'],
            'YTD_BUDGET':        ytd_budget,
            'YTD_ACTUAL':        ytd_actual,
            'YTD_VARIANCE':      ytd_variance,
            'AggregationMethod': 'SUM'
        })
        print(f"  [SUM] {data_point}: Budget={ytd_budget:.2f}, Actual={ytd_actual:.2f}, Variance={ytd_variance:.2f}")

    # Step 1b: Derive TDI and Adjusted Indirect Costs (needed before % derivation)
    abs_lookup = {r['DataPoint']: r for r in consolidated}

    def _get(key, field):
        m = abs_lookup.get(key)
        return m[field] if m else 0.0

    tdi_b = sum(_get(c, 'YTD_BUDGET')   for c in TDI_COMPONENTS)
    tdi_a = sum(_get(c, 'YTD_ACTUAL')   for c in TDI_COMPONENTS)
    tdi_v = sum(_get(c, 'YTD_VARIANCE') for c in TDI_COMPONENTS)

    tdi_row = {
        'DataPoint': 'TDI', 'DashboardName': 'Tax, Depreciation & Interest',
        'YTD_BUDGET': tdi_b, 'YTD_ACTUAL': tdi_a, 'YTD_VARIANCE': tdi_v,
        'AggregationMethod': 'DERIVED'
    }
    consolidated.append(tdi_row)
    abs_lookup['TDI'] = tdi_row
    print(f"  [DERIVED] TDI: Budget={tdi_b:.2f}, Actual={tdi_a:.2f}, Variance={tdi_v:.2f}")

    adj_b = _get('Total Fixed Cost (Indirect)', 'YTD_BUDGET')   - tdi_b
    adj_a = _get('Total Fixed Cost (Indirect)', 'YTD_ACTUAL')   - tdi_a
    adj_v = _get('Total Fixed Cost (Indirect)', 'YTD_VARIANCE') - tdi_v

    adj_row = {
        'DataPoint': 'Adjusted Indirect Costs', 'DashboardName': 'Indirect Costs',
        'YTD_BUDGET': adj_b, 'YTD_ACTUAL': adj_a, 'YTD_VARIANCE': adj_v,
        'AggregationMethod': 'DERIVED'
    }
    consolidated.append(adj_row)
    abs_lookup['Adjusted Indirect Costs'] = adj_row
    print(f"  [DERIVED] Adjusted Indirect Costs: Budget={adj_b:.2f}, Actual={adj_a:.2f}, Variance={adj_v:.2f}")

    # Step 2: Derive percentage metrics from YTD absolutes
    # abs_lookup already populated above

    for pct_metric, (num_key, den_key, sign) in PCT_DERIVATIONS.items():
        if pct_metric not in grouped:
            continue

        num = abs_lookup.get(num_key)
        den = abs_lookup.get(den_key)
        dashboard_name = grouped[pct_metric]['dashboard_name']

        pct_budget = num['YTD_BUDGET'] / den['YTD_BUDGET'] if (num and den and den['YTD_BUDGET']) else 0.0
        pct_actual = num['YTD_ACTUAL'] / den['YTD_ACTUAL'] if (num and den and den['YTD_ACTUAL']) else 0.0
        pct_variance = (pct_budget - pct_actual) if sign == 'cost' else (pct_actual - pct_budget)

        consolidated.append({
            'DataPoint':         pct_metric,
            'DashboardName':     dashboard_name,
            'YTD_BUDGET':        pct_budget,
            'YTD_ACTUAL':        pct_actual,
            'YTD_VARIANCE':      pct_variance,
            'AggregationMethod': 'DERIVED'
        })
        print(f"  [DERIVED] {pct_metric}: Budget={pct_budget*100:.3f}%, Actual={pct_actual*100:.3f}%, Variance={pct_variance*100:.3f}%")

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
        '--company',
        required=True,
        help='Company ID (e.g. mudin, mantis)'
    )
    parser.add_argument(
        "--input",
        default=None,
        help="Path to raw data JSON file"
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Path to output consolidated JSON file"
    )

    args = parser.parse_args()

    if args.input is None:
        args.input = f"dashboard/data/raw-data-{args.company}.json"
    if args.output is None:
        args.output = f"dashboard/data/calculations-{args.company}.json"

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

        # Derive EBITDA from components
        def get_metric(data, name, field):
            m = next((r for r in data if r['DataPoint'] == name), None)
            return m[field] if m else 0.0

        ebitda_components = ['Net Surplus / (Deflect)', 'Interest Expenses', 'Amortization', 'Depreciation']
        revenue_budget = get_metric(consolidated, 'Total Revenue', 'YTD_BUDGET')
        revenue_actual = get_metric(consolidated, 'Total Revenue', 'YTD_ACTUAL')

        ebitda_budget = sum(get_metric(consolidated, c, 'YTD_BUDGET') for c in ebitda_components)
        ebitda_actual = sum(get_metric(consolidated, c, 'YTD_ACTUAL') for c in ebitda_components)
        ebitda_variance = ebitda_actual - ebitda_budget

        ebitda_pct_budget = ebitda_budget / revenue_budget if revenue_budget else 0.0
        ebitda_pct_actual = ebitda_actual / revenue_actual if revenue_actual else 0.0
        ebitda_pct_variance = ebitda_pct_actual - ebitda_pct_budget

        consolidated.append({
            'DataPoint': 'EBITDA',
            'DashboardName': 'EBITDA',
            'YTD_BUDGET': ebitda_budget,
            'YTD_ACTUAL': ebitda_actual,
            'YTD_VARIANCE': ebitda_variance,
            'AggregationMethod': 'DERIVED'
        })
        consolidated.append({
            'DataPoint': 'EBITDA %',
            'DashboardName': 'EBITDA %',
            'YTD_BUDGET': ebitda_pct_budget,
            'YTD_ACTUAL': ebitda_pct_actual,
            'YTD_VARIANCE': ebitda_pct_variance,
            'AggregationMethod': 'DERIVED'
        })

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
