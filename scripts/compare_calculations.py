"""
Budget vs Actual - Validation Script

Compares Excel-based calculations vs JSON-based calculations.
Ensures both independent calculation paths produce identical results.
Acts as data quality check before dashboard deployment.
"""

import argparse
import json
import openpyxl
import sys
from pathlib import Path


# Absolute metrics to validate (percentage metrics excluded — they're computed as
# YTD numerator/denominator in JSON vs monthly AVERAGEIF in Excel, so they differ by design)
METRICS_TO_VALIDATE = [
    "Total Revenue",
    "Total Variable Cost",
    "Contribution Margin",
    "Total Staff Cost (Direct)",
    "Gross Profit / (Loss)",
    "Total Fixed Cost (Indirect)",
    "Net Surplus / (Deflect)",
    "Interest Expenses",
    "Amortization",
    "Depreciation",
]


def read_excel_calculations(excel_path="Financial-Data-Database.xlsx"):
    """
    Compute YTD sums from BVA_DATA sheet (SUMIF equivalent in Python).
    This avoids dependency on Excel formula cache, which openpyxl cannot populate.

    Returns:
        dict: {metric_name: {'budget': value, 'actual': value, 'variance': value}}
    """
    print(f"Reading Excel calculations from {excel_path}...")

    wb = openpyxl.load_workbook(excel_path, data_only=True)

    if 'BVA_DATA' not in wb.sheetnames:
        print(f"ERROR: BVA_DATA sheet not found in {excel_path}")
        wb.close()
        return None

    ws = wb['BVA_DATA']

    # BVA_DATA columns: Month | DataPoint | SampleRow | DashboardName | Budget | Actual | Variance
    # Find header row (row 1)
    totals = {}

    for row in range(2, ws.max_row + 1):
        metric = ws.cell(row, 2).value  # Column B = DataPoint
        if not metric or metric not in METRICS_TO_VALIDATE:
            continue

        budget   = ws.cell(row, 5).value  # Column E
        actual   = ws.cell(row, 6).value  # Column F
        variance = ws.cell(row, 7).value  # Column G

        budget   = float(budget)   if budget   is not None else 0.0
        actual   = float(actual)   if actual   is not None else 0.0
        variance = float(variance) if variance is not None else 0.0

        if metric not in totals:
            totals[metric] = {'budget': 0.0, 'actual': 0.0, 'variance': 0.0}
        totals[metric]['budget']   += budget
        totals[metric]['actual']   += actual
        totals[metric]['variance'] += variance

    # Round to 2dp to match JSON
    for m in totals:
        totals[m] = {k: round(v, 2) for k, v in totals[m].items()}

    wb.close()

    print(f"  Found {len(totals)} metrics in Excel")
    return totals


def read_json_calculations(json_path="dashboard/data/calculations.json"):
    """
    Read YTD calculations from JSON file.

    Returns:
        dict: {metric_name: {'budget': value, 'actual': value, 'variance': value}}
    """
    print(f"Reading JSON calculations from {json_path}...")

    json_path = Path(json_path)

    if not json_path.exists():
        print(f"ERROR: {json_path} not found")
        return None

    with open(json_path, 'r') as f:
        data = json.load(f)

    results = {}

    for item in data:
        metric = item.get('DataPoint')

        if metric and metric in METRICS_TO_VALIDATE:
            results[metric] = {
                'budget': item.get('YTD_BUDGET', 0.0),
                'actual': item.get('YTD_ACTUAL', 0.0),
                'variance': item.get('YTD_VARIANCE', 0.0)
            }

    print(f"  Found {len(results)} metrics in JSON")
    return results


def compare_values(excel_val, json_val, tolerance=0.01):
    """
    Compare two values with tolerance for floating point precision.

    Args:
        excel_val: Value from Excel
        json_val: Value from JSON
        tolerance: Maximum acceptable difference

    Returns:
        tuple: (match: bool, difference: float)
    """
    diff = abs(excel_val - json_val)
    match = diff <= tolerance
    return match, diff


def validate_calculations(excel_results, json_results):
    """
    Compare Excel and JSON calculations for all metrics.

    Returns:
        bool: True if all match, False if any differences found
    """
    print("\n" + "=" * 70)
    print("VALIDATION: Comparing Excel vs JSON Calculations")
    print("=" * 70)
    print()

    all_match = True
    differences = []

    print(f"Checking {len(METRICS_TO_VALIDATE)} metrics:")
    print()

    for metric in METRICS_TO_VALIDATE:
        # Check if metric exists in both
        if metric not in excel_results:
            print(f"  [MISS] {metric:35s} - Missing in Excel")
            all_match = False
            continue

        if metric not in json_results:
            print(f"  [MISS] {metric:35s} - Missing in JSON")
            all_match = False
            continue

        excel_data = excel_results[metric]
        json_data = json_results[metric]

        # Compare Budget, Actual, Variance
        budget_match, budget_diff = compare_values(excel_data['budget'], json_data['budget'])
        actual_match, actual_diff = compare_values(excel_data['actual'], json_data['actual'])
        variance_match, variance_diff = compare_values(excel_data['variance'], json_data['variance'])

        if budget_match and actual_match and variance_match:
            print(f"  [OK]   {metric:35s} Excel={excel_data['actual']:>15,.2f}  JSON={json_data['actual']:>15,.2f}")
        else:
            print(f"  [FAIL] {metric:35s} MISMATCH:")
            if not budget_match:
                print(f"     Budget:   Excel={excel_data['budget']:>15,.2f}  JSON={json_data['budget']:>15,.2f}  Diff={budget_diff:>15,.2f}")
            if not actual_match:
                print(f"     Actual:   Excel={excel_data['actual']:>15,.2f}  JSON={json_data['actual']:>15,.2f}  Diff={actual_diff:>15,.2f}")
            if not variance_match:
                print(f"     Variance: Excel={excel_data['variance']:>15,.2f}  JSON={json_data['variance']:>15,.2f}  Diff={variance_diff:>15,.2f}")

            differences.append({
                'metric': metric,
                'excel': excel_data,
                'json': json_data,
                'budget_diff': budget_diff,
                'actual_diff': actual_diff,
                'variance_diff': variance_diff
            })
            all_match = False

    print()
    print("=" * 70)

    if all_match:
        print("[PASS] VALIDATION PASSED!")
        print("All Excel and JSON calculations match perfectly.")
        print("Proceeding with dashboard generation...")
        print("=" * 70)
        return True
    else:
        print("[FAIL] VALIDATION FAILED!")
        print()
        print(f"{len(differences)} metric(s) don't match:")
        for diff in differences:
            print(f"  - {diff['metric']}")
        print()
        print("INVESTIGATION REQUIRED:")
        print("  1. Check BVA_DATA sheet in Financial-Data-Database.xlsx")
        print("  2. Check dashboard/data/raw-data.json")
        print("  3. Verify source data extraction")
        print("  4. Check formulas in BVA_CALC sheet")
        print()
        print("Process STOPPED. Fix issues before deploying.")
        print("=" * 70)
        return False


def main():
    """
    Main validation workflow.
    """
    parser = argparse.ArgumentParser(
        description="Compare Excel vs JSON calculations"
    )
    parser.add_argument('--company', required=True, help='Company ID (e.g. mudin, mantis)')
    args = parser.parse_args()
    excel_path = f"Financial-Data-Database-{args.company.capitalize()}.xlsx"
    json_path = f"dashboard/data/calculations-{args.company}.json"

    print("=" * 70)
    print("BUDGET VS ACTUAL - CALCULATION VALIDATION")
    print("=" * 70)
    print()

    # Read Excel calculations
    excel_results = read_excel_calculations(excel_path)
    if excel_results is None:
        print("\nERROR: Could not read Excel calculations")
        sys.exit(1)

    # Read JSON calculations
    json_results = read_json_calculations(json_path)
    if json_results is None:
        print("\nERROR: Could not read JSON calculations")
        sys.exit(1)

    # Validate
    validation_passed = validate_calculations(excel_results, json_results)

    # Exit with appropriate code
    if validation_passed:
        sys.exit(0)  # Success
    else:
        sys.exit(1)  # Failure


if __name__ == "__main__":
    main()
