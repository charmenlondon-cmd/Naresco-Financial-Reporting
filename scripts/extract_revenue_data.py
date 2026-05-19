"""
Revenue Data Extraction Script

Extracts contract revenue line items from Financial Statement source files (REVENUE tab).
Writes/updates a Revenue sheet in the company database with month-by-month columns
and a running YTD Total that always sits to the right of the latest month.
"""

import openpyxl
from openpyxl import Workbook, load_workbook
import argparse
from pathlib import Path
from datetime import datetime


MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
]

HEADER_ROW = 6
SOURCE_SHEET = "REVENUE"
DB_SHEET = "Revenue"
TOTAL_LABEL = "TOTAL CONTRACT REVENUE"


def month_label(dt):
    """Convert datetime to label e.g. 'Jan-26'."""
    return f"{MONTH_NAMES[dt.month - 1][:3]}-{str(dt.year)[2:]}"


def detect_month_columns(ws):
    """
    Find columns in header row that contain datetime month values.
    Returns list of (col_index, datetime, label) sorted by date, skipping 2025 and Total columns.
    """
    months = []
    for col in range(3, ws.max_column + 1):
        v = ws.cell(HEADER_ROW, col).value
        if isinstance(v, datetime):
            months.append((col, v, month_label(v)))
    return sorted(months, key=lambda x: x[1])


def extract_source_data(ws, month_cols):
    """
    Extract contract revenue rows and section totals from source REVENUE tab.

    Returns:
        rows: list of dicts {code, particulars, 'Jan-26': val, 'Feb-26': val, ...}
        totals: dict {'Jan-26': val, 'Feb-26': val, ...}
    """
    rows = []
    totals = {}

    for row in range(HEADER_ROW + 1, ws.max_row + 1):
        col_c = ws.cell(row, 3).value  # CODE
        col_d = ws.cell(row, 4).value  # PARTICULARS

        # Stop when "Other Revenue" section is reached
        if isinstance(col_d, str) and "other revenue" in col_d.lower():
            break

        # Totals row: no CODE, no PARTICULARS, but has month values
        if col_c is None and col_d is None:
            row_vals = {}
            for src_col, dt, label in month_cols:
                v = ws.cell(row, src_col).value
                if isinstance(v, (int, float)) and v != 0:
                    row_vals[label] = round(float(v), 2)
            if row_vals:
                totals = row_vals
            continue

        # Skip rows without a numeric CODE or without PARTICULARS
        if not isinstance(col_c, (int, float)):
            continue
        if not col_d or str(col_d).strip() == "":
            continue

        row_data = {
            "code": int(col_c),
            "particulars": str(col_d).strip() if col_d else ""
        }
        for src_col, dt, label in month_cols:
            v = ws.cell(row, src_col).value
            row_data[label] = round(float(v), 2) if isinstance(v, (int, float)) else 0.0
        rows.append(row_data)

    return rows, totals


def read_existing_db(ws):
    """
    Read existing Revenue sheet from database.

    Returns:
        existing_months: list of month labels already in database (in order)
        existing_rows: dict keyed by (code, particulars) -> {label: value, ...}
        existing_totals: dict {label: value}
    """
    if ws.max_row < 1:
        return [], {}, {}

    # Read header row (row 1)
    existing_months = []
    ytd_col = None
    col = 3
    while True:
        v = ws.cell(1, col).value
        if v is None:
            break
        if v == "YTD Total":
            ytd_col = col
            break
        existing_months.append((col, v))
        col += 1

    # Read data rows
    existing_rows = {}
    existing_totals = {}

    for row in range(2, ws.max_row + 1):
        code = ws.cell(row, 1).value
        particulars = ws.cell(row, 2).value

        if particulars == TOTAL_LABEL:
            for db_col, label in existing_months:
                v = ws.cell(row, db_col).value
                existing_totals[label] = round(float(v), 2) if isinstance(v, (int, float)) else 0.0
            continue

        if code is None or particulars is None:
            continue

        key = (int(code), str(particulars).strip())
        row_vals = {}
        for db_col, label in existing_months:
            v = ws.cell(row, db_col).value
            row_vals[label] = round(float(v), 2) if isinstance(v, (int, float)) else 0.0
        existing_rows[key] = row_vals

    return existing_months, existing_rows, existing_totals


def write_revenue_sheet(ws, all_months, all_rows_ordered, all_totals):
    """
    Write/rewrite the Revenue sheet with all month columns + YTD Total.

    Args:
        ws: worksheet (will be cleared and rewritten)
        all_months: list of month labels in order e.g. ['Jan-26', 'Feb-26']
        all_rows_ordered: list of dicts {code, particulars, 'Jan-26': val, ...}
        all_totals: dict {label: value}
    """
    ws.delete_rows(1, ws.max_row)

    # Write header row
    ws.cell(1, 1, "CODE")
    ws.cell(1, 2, "PARTICULARS")
    for i, label in enumerate(all_months):
        ws.cell(1, 3 + i, label)
    ytd_col = 3 + len(all_months)
    ws.cell(1, ytd_col, "YTD Total")

    # Write data rows
    for r_idx, row_data in enumerate(all_rows_ordered, start=2):
        ws.cell(r_idx, 1, row_data["code"])
        ws.cell(r_idx, 2, row_data["particulars"])
        row_total = 0.0
        for i, label in enumerate(all_months):
            v = row_data.get(label, 0.0)
            ws.cell(r_idx, 3 + i, v)
            row_total += v
        ws.cell(r_idx, ytd_col, round(row_total, 2))

    # Write totals row
    totals_row = 2 + len(all_rows_ordered)
    ws.cell(totals_row, 1, "")
    ws.cell(totals_row, 2, TOTAL_LABEL)
    ytd_total = 0.0
    for i, label in enumerate(all_months):
        v = all_totals.get(label, 0.0)
        ws.cell(totals_row, 3 + i, v)
        ytd_total += v
    ws.cell(totals_row, ytd_col, round(ytd_total, 2))

    print(f"  Written {len(all_rows_ordered)} line items + totals row")
    print(f"  Months: {', '.join(all_months)}")


def update_database(database_path, source_rows, source_totals, source_months):
    """
    Open/create database and update the Revenue sheet with new month data.
    Only adds months not already present in the database.
    """
    database_path = Path(database_path)

    if database_path.exists():
        wb = load_workbook(database_path)
        print(f"Opened existing database: {database_path}")
    else:
        wb = Workbook()
        wb.active.title = "Sheet1"
        print(f"Creating new database: {database_path}")

    # Get or create Revenue sheet
    if DB_SHEET in wb.sheetnames:
        ws = wb[DB_SHEET]
        existing_months_tuples, existing_rows, existing_totals = read_existing_db(ws)
        existing_month_labels = [label for _, label in existing_months_tuples]
        print(f"  Existing months in database: {existing_month_labels or 'none'}")
    else:
        ws = wb.create_sheet(DB_SHEET)
        existing_month_labels = []
        existing_rows = {}
        existing_totals = {}
        print(f"  Creating new {DB_SHEET} sheet")

    # Determine which months are new
    new_months = [label for label in source_months if label not in existing_month_labels]
    if not new_months:
        print(f"  No new months to add. Database already contains: {existing_month_labels}")
        wb.close()
        return False

    print(f"  Adding new month(s): {new_months}")

    # Build merged month list (existing + new, in chronological order)
    all_months = existing_month_labels + new_months

    # Build merged row data — existing rows updated with new month values
    # Key: (code, particulars)
    source_by_key = {(r["code"], r["particulars"]): r for r in source_rows}

    # Start from existing rows, preserving order
    merged_rows = {}
    for key, vals in existing_rows.items():
        merged_rows[key] = dict(vals)

    # Add new month values to existing rows, and add any brand-new rows
    for key, src_data in source_by_key.items():
        if key not in merged_rows:
            merged_rows[key] = {label: 0.0 for label in existing_month_labels}
        for label in new_months:
            merged_rows[key][label] = src_data.get(label, 0.0)

    # For any existing rows NOT in source this month, set new month values to 0
    for key in merged_rows:
        for label in new_months:
            if label not in merged_rows[key]:
                merged_rows[key][label] = 0.0

    # Build ordered rows list (preserving insertion order)
    all_rows_ordered = []
    for key, vals in merged_rows.items():
        all_rows_ordered.append({
            "code": key[0],
            "particulars": key[1],
            **vals
        })

    # Merge totals
    for label in new_months:
        existing_totals[label] = source_totals.get(label, 0.0)

    # Write the sheet
    write_revenue_sheet(ws, all_months, all_rows_ordered, existing_totals)

    wb.save(database_path)
    wb.close()
    print(f"  Database saved: {database_path}")
    return True


def main():
    parser = argparse.ArgumentParser(
        description="Extract revenue line items from Financial Statement and update company database"
    )
    parser.add_argument("--input", required=True, help="Path to source Financial Statement Excel file")
    parser.add_argument("--company", required=True, help="Company ID (e.g. mudin, mantis)")
    parser.add_argument("--database", default=None, help="Path to company database (default: Financial-Data-Database-{Company}.xlsx)")
    parser.add_argument("--sheet", default=SOURCE_SHEET, help=f"Source sheet name (default: {SOURCE_SHEET})")

    args = parser.parse_args()

    if args.database is None:
        args.database = f"Financial-Data-Database-{args.company.capitalize()}.xlsx"

    print("=" * 60)
    print("REVENUE DATA EXTRACTION")
    print("=" * 60)
    print(f"Source:   {args.input}")
    print(f"Database: {args.database}")
    print(f"Company:  {args.company}")
    print()

    try:
        wb = openpyxl.load_workbook(args.input, data_only=True)
        if args.sheet not in wb.sheetnames:
            raise ValueError(f"Sheet '{args.sheet}' not found. Available: {wb.sheetnames}")
        ws = wb[args.sheet]

        # Detect month columns
        month_cols = detect_month_columns(ws)
        if not month_cols:
            raise ValueError("No month columns (datetime headers) found in header row.")
        print(f"Found months in source: {[label for _, _, label in month_cols]}")

        # Extract data
        source_rows, source_totals = extract_source_data(ws, month_cols)
        source_month_labels = [label for _, _, label in month_cols]

        print(f"Extracted {len(source_rows)} line items")
        print(f"Section totals: {source_totals}")
        wb.close()

        # Update database
        update_database(args.database, source_rows, source_totals, source_month_labels)

        print()
        print("=" * 60)
        print("[SUCCESS] REVENUE EXTRACTION COMPLETE")
        print("=" * 60)

    except Exception as e:
        print(f"\n[ERROR] {e}")
        raise


if __name__ == "__main__":
    main()
