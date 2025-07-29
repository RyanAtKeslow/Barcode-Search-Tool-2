#!/usr/bin/env python3
"""
format_lens_sheet.py
--------------------
Flatten a complex lens inventory CSV (e.g. "Copy of Keslow Lens Inventory - Full Frame Spherical.csv")
into a single, tidy table with one header row and 19 fixed columns.

The source file repeats a mini-header (",,Focal Length,T-Stop, …") in between lens sets and
leaves manufacturer/series cells blank on subsequent rows within the same set.

This script:
1. Scans each row and discards any that look like a header (contain "Focal Length").
2. Propagates the last seen Manufacturer & Series when a row leaves them blank.
3. Pads/trim rows to exactly 19 columns (A-S) using the schema below.
4. Outputs `Flattened_Lens_Inventory.csv` next to the script.

Columns (A-S):
 0  Manufacturer
 1  Series
 2  Focal Length
 3  T-Stop
 4  Close Focus
 5  Weight (lbs)
 6  Length (in)
 7  Front Diameter (mm)
 8  Mount
 9  Image Circle (mm)
10  Film Compatibility
11  Iris Blade Count
12  Extender
13  LDS
14  Support Post Length (mm)
15  Manufacture Year
16  Notes
17  (reserved / blank)
18  (reserved / blank)

You can rename / reorder the last two columns later if needed – they are kept so
we always reach column S.

Usage:
    python3 format_lens_sheet.py SOURCE_CSV_PATH

If no path is supplied the script assumes the file lives beside it under the
name "Copy of Keslow Lens Inventory - Full Frame Spherical.csv".
"""
import csv
import sys
from pathlib import Path

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
DEFAULT_SOURCE = Path(__file__).with_name("Copy of Keslow Lens Inventory - Sheet3")
OUTPUT_FILE    = Path(__file__).with_name("Flattened_Lens_Inventory.csv")

HEADERS = [
    'Manufacturer', 'Series', 'Focal Length', 'T-Stop', 'Close Focus',
    'Weight (lbs)', 'Length (in)', 'Front Diameter (mm)', 'Mount',
    'Image Circle (mm)', 'Film Compatibility', 'Iris Blade Count',
    'Extender', 'LDS', 'Support Post Length (mm)', 'Manufacture Year',
    'Notes', 'Extra 1', 'Extra 2'
]
TOTAL_COLS = len(HEADERS)  # 19

# ----------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------

def looks_like_section_header(row: list[str]) -> bool:
    """Return True if the row is the mini-header that precedes each lens set."""
    return any(cell.strip().lower() == 'focal length' for cell in row)


def pad(row: list[str], size: int) -> list[str]:
    """Pad or trim the row to exactly *size* elements."""
    row = row[:size]
    if len(row) < size:
        row += [''] * (size - len(row))
    return row

# ----------------------------------------------------------------------------
# Main flatten routine
# ----------------------------------------------------------------------------

def flatten(source: Path, dest: Path):
    if not source.exists():
        sys.exit(f"Source file not found: {source}")

    output_rows = []
    current_manufacturer = ''
    current_series = ''

    with source.open(newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        for raw in reader:
            if not any(raw):
                continue  # skip completely blank lines
            if looks_like_section_header(raw):
                # This row only defines column names for the following block.
                continue

            # Propagate manufacturer & series from previous row if omitted
            manufacturer = raw[0].strip()
            series       = raw[1].strip() if len(raw) > 1 else ''

            if manufacturer:
                current_manufacturer = manufacturer
            else:
                manufacturer = current_manufacturer

            if series:
                current_series = series
            else:
                series = current_series

            # Build row up to TOTAL_COLS
            full_row = pad([manufacturer, series] + raw[2:], TOTAL_COLS)
            output_rows.append(dict(zip(HEADERS, full_row)))

    # Write result
    with dest.open('w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=HEADERS)
        writer.writeheader()
        writer.writerows(output_rows)

    print(f"Flattened {len(output_rows)} rows → {dest}")


if __name__ == '__main__':
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_SOURCE
    flatten(src, OUTPUT_FILE) 