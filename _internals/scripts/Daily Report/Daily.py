#!/usr/bin/env python3
"""
Daily Report Generator
======================
Generates daily LED change out reports from Excel inventory files.

This script:
1. Reads from Excel files in the Input folder
2. Processes date information
3. Groups data by date and week
4. Outputs formatted text reports

Configuration:
- Edit config.py to customize input/output folders and settings
- Or use command-line defaults (prompts for region and week count)

Usage:
    python Daily.py                           # Interactive mode
    python Daily.py --region "Ventura"        # With region specified
    python Daily.py --weeks 4                 # For 4 weeks
"""

import pandas as pd
import sys
import os
from pathlib import Path
from datetime import datetime, timedelta, date
from types import SimpleNamespace

# ============================================================================
# Configuration & Paths
# ============================================================================

# Get the script's directory
SCRIPT_DIR = Path(__file__).parent.resolve()

# Try to load from config.py if available
try:
    sys.path.insert(0, str(Path(__file__).parent.parent.parent / "config"))
    import config

    INPUT_DIR = Path(config.DAILY_REPORT_SETTINGS.get("input_folder"))
    OUTPUT_DIR = Path(config.DAILY_REPORT_SETTINGS.get("output_folder"))
except (ImportError, KeyError, AttributeError):
    # Fallback to script-relative directories if config not available
    INPUT_DIR = SCRIPT_DIR / "Input"
    OUTPUT_DIR = SCRIPT_DIR / "Output"

# Create directories if they don't exist
INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Default configuration
DEFAULT_EXCEL = str(INPUT_DIR)
DEFAULT_SHEET = None              # Use first sheet by default
DEFAULT_DATE_COL = "LED Change Date"
DEFAULT_PATTERN = "*.xls*"         # Matches .xlsx, .xlsm, .xls

# ============================================================================
# Helper Functions
# ============================================================================

def find_latest_file(folder: Path, pattern: str) -> Path:
    """Find the most recently modified file matching the pattern"""
    files = [f for f in folder.glob(pattern) if f.is_file() and not f.name.startswith("~$")]
    if not files:
        raise FileNotFoundError(f"No files found in {folder} matching pattern: {pattern}")
    files.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return files[0]

def load_df(path_like, sheet, pattern):
    """Load Excel file into a DataFrame"""
    path = Path(path_like)
    if path.is_dir():
        path = find_latest_file(path, pattern)
        print(f"[Auto-selected latest file: {path.name}]")

    if sheet:
        return pd.read_excel(path, sheet_name=sheet)
    return pd.read_excel(path)

def coerce_date_col(df, date_col):
    """Convert a column to dates and remove invalid entries"""
    if date_col not in df.columns:
        raise ValueError(f"Date column '{date_col}' not found. Available columns: {list(df.columns)}")

    d = df.copy()
    d[date_col] = pd.to_datetime(d[date_col], errors="coerce").dt.date
    d = d.dropna(subset=[date_col])
    return d

def apply_optional_filter(df, col, val):
    """Apply an optional filter to the DataFrame"""
    if col is None or val is None:
        return df
    if col not in df.columns:
        raise ValueError(f"Filter column '{col}' not found. Available columns: {list(df.columns)}")
    return df.loc[df[col] == val].copy()

def monday_of_week(d: date) -> date:
    """Get the Monday of the week containing date d"""
    return d - timedelta(days=d.weekday())

def week_span(asof: date, n_back: int):
    """Get the Monday-Sunday span for n weeks back from asof date"""
    this_mon = monday_of_week(asof)
    wk_mon = this_mon - timedelta(days=7 * n_back)
    wk_sun = wk_mon + timedelta(days=6)
    return wk_mon, wk_sun

def build_daily_counts(df, date_col):
    """Build a daily summary with running totals"""
    daily = df.groupby(date_col, dropna=False).size().rename("Total").reset_index()
    daily = daily.sort_values(date_col).reset_index(drop=True)
    daily["Total to Date"] = daily["Total"].cumsum()
    return daily

def format_mdyy(d: date) -> str:
    """Format date as M.D.YY"""
    yy = str(d.year)[-2:]
    return f"{d.month}.{d.day}.{yy}"

def generate_week_block(daily_all, date_col, week_start, week_end, region, include_zero=False):
    """Generate a formatted report block for a specific week"""
    day_map = {row[date_col]: (int(row["Total"]), int(row["Total to Date"]))
               for _, row in daily_all.iterrows()}

    days = [week_start + timedelta(days=i) for i in range(7)]
    records = []
    last_cum = 0

    # Find the last cumulative value before this week
    for day in sorted(day_map.keys()):
        if day <= week_start:
            last_cum = day_map[day][1]

    # Build records for each day in the week
    for day in days:
        total, cum = day_map.get(day, (0, None))
        if cum is None:
            cum = last_cum
        else:
            last_cum = cum

        # Skip days with no activity unless include_zero is True
        if total == 0 and not include_zero:
            continue

        date_str = format_mdyy(day)
        body = (
            f"Daily Report {region} {date_str}\n"
            f"Total: {total}\n"
            f"Total to Date: {cum}"
        )
        records.append((day, body))

    # Sort by date (newest first)
    records.sort(key=lambda r: r[0], reverse=True)

    # Format header
    header = f"=== Week {format_mdyy(week_start)} – {format_mdyy(week_end)} ==="

    if records:
        return header + "\n\n" + "\n\n".join(b for _, b in records)
    else:
        return header + ("\n(no activity this week)" if include_zero else "\n(no reportable days)")

def auto_output_path(excel_path: str, earliest_start: date, latest_end: date, region: str) -> Path:
    """Generate an output filename based on date range and region"""
    start_label = earliest_start.strftime("%Y-%m-%d")
    end_label = latest_end.strftime("%Y-%m-%d")
    region_part = f"{region}_" if region else ""
    out_name = f"Daily_Report_{region_part}{start_label}_to_{end_label}.txt"
    return OUTPUT_DIR / out_name

# ============================================================================
# Main Function
# ============================================================================

def main():
    """Main report generation logic"""

    # Prompt for region name (blank if user just presses Enter)
    region = input('Enter the city/region for the report (or press Enter for none): ').strip()

    # Prompt for number of weeks
    weeks_input = input('Enter number of weeks to include (default: 2): ').strip() or '2'
    try:
        weeks = int(weeks_input)
    except ValueError:
        weeks = 2

    # Build configuration namespace
    args = SimpleNamespace(
        excel=str(INPUT_DIR),
        pattern=DEFAULT_PATTERN,
        sheet=DEFAULT_SHEET,
        date_col=DEFAULT_DATE_COL,
        asof=None,
        region=region,
        filter_col=None,
        filter_val=None,
        include_zero=False,
        outtxt=None,
        weeks=weeks,
    )

    # Validate weeks
    asof = date.today() if args.asof is None else datetime.strptime(args.asof, "%Y-%m-%d").date()
    if args.weeks < 1:
        raise ValueError("Number of weeks must be at least 1")

    # Load and process data
    try:
        print("\n📂 Looking for Excel files in:", INPUT_DIR)
        df = load_df(args.excel, args.sheet, args.pattern)
    except FileNotFoundError as e:
        print(f"\n❌ Error: {e}")
        print(f"\nPlease add Excel files to: {INPUT_DIR}")
        return

    # Apply filters and process dates
    df = apply_optional_filter(df, args.filter_col, args.filter_val)

    try:
        df = coerce_date_col(df, args.date_col)
    except ValueError as e:
        print(f"\n❌ Error: {e}")
        return

    if df.empty:
        print("\n❌ No rows with valid dates found after filtering.")
        return

    # Generate daily counts
    daily_all = build_daily_counts(df, args.date_col)

    # Generate report blocks for each week
    blocks = []
    week_spans = []
    for n in range(1, args.weeks + 1):
        wk_start, wk_end = week_span(asof, n)
        week_spans.append((wk_start, wk_end))
        block = generate_week_block(
            daily_all=daily_all,
            date_col=args.date_col,
            week_start=wk_start,
            week_end=wk_end,
            region=args.region,
            include_zero=args.include_zero,
        )
        blocks.append(block)

    # Combine blocks
    earliest_start = min(ws for ws, _ in week_spans)
    latest_end = max(we for _, we in week_spans)

    body = "\n\n" + ("-" * 60) + "\n\n"
    body = body.join(blocks)

    # Save report
    out_path = Path(args.outtxt) if args.outtxt else auto_output_path(
        args.excel, earliest_start, latest_end, args.region
    )

    try:
        out_path.write_text(body, encoding="utf-8")
        print("\n" + body)
        print(f"\n✅ Report saved: {out_path}")
    except IOError as e:
        print(f"\n❌ Error saving report: {e}")

# ============================================================================
# Entry Point
# ============================================================================

if __name__ == "__main__":
    try:
        main()
    except BrokenPipeError:
        # Handle broken pipe gracefully
        pass
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        sys.exit(1)
