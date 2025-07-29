#!/usr/bin/env python3
"""
format_lens_sheet.py
--------------------
Aggregate/flatten every *.csv found in a target directory (default: "To Be Parsed")
into a single table with a canonical header used by the Lens Database project.

Enhancements:
• Sets "Prime / Zoom / Special" to "Zoom" if Focal Length contains a dash (-),
  otherwise "Prime".
• If the word "Macro" appears in the Focal Length, it is removed from that
  cell and appended to the Notes column.
• Removes any variant of "not in inventory" from the Notes cell.

Usage:
    # Use default folder
    python3 format_lens_sheet.py

    # Specify another folder
    python3 format_lens_sheet.py /path/to/folder

Outputs a single CSV called `Flattened_Lens_Inventory.csv` in the script's
directory.
"""
import csv
import sys
import re
from pathlib import Path
from typing import List, Dict, Optional

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
PROJECT_DIR = Path(__file__).parent
DEFAULT_FOLDER = PROJECT_DIR / "To Be Parsed"
OUTPUT_FILE = PROJECT_DIR / "Flattened_Lens_Inventory.csv"

HEADERS = [
    'Manufacturer', 'Series', 'Focal Length', 'T-Stop',
    'Prime / Zoom / Special', 'Format', 'Mount', 'Anamorphic / Spherical',
    'Anamorphic Squeeze Factor', 'Anamorphic Location', 'Housing',
    'Front Diameter (mm)', 'Close Focus', 'Length (in)', 'Film Compatibility',
    'Image Circle (mm)', 'Iris Blade Count', 'Extender', 'LDS', 'i/Data',
    'Support Recommended', 'Support Post Length (mm)', 'Weight (lbs)',
    'Manufacture Year', 'Expander', 'Heden Motor Size', 'Size', 'Notes',
    'Look', 'Use Case', 'Bokeh', 'Flare', 'Focus Falloff', 'Breathing',
    'Focus Scale'
]
TOTAL_COLS = len(HEADERS)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def looks_like_section_header(row: List[str]) -> bool:
    return any('focal' in cell.lower() and 'length' in cell.lower() for cell in row)


# Normalise header strings for matching
def norm(s: str) -> str:
    return re.sub(r'[^a-z0-9]', '', s.lower())


# Build mapping from column index to master header
def build_header_map(row: List[str]) -> Dict[int, str]:
    mapping: Dict[int, str] = {}
    alias_map = {
        # T-Stop variations
        'tstop': 'T-Stop',
        't': 'T-Stop',
        'minstop': 'T-Stop',
        'maxstop': 'T-Stop',
        # Focal length synonyms already match
        # Weight column
        'weightlbs': 'Weight (lbs)',
        'weight': 'Weight (lbs)',
        # Length column
        'lengthin': 'Length (in)',
        # Front diameter
        'frontdiameter': 'Front Diameter (mm)',
        'frontdiametermm': 'Front Diameter (mm)',
        # Image circle
        'imagecirclemm': 'Image Circle (mm)',
        # format alias such as imagecircle, format etc handled
    }

    normalized_master = {norm(h): h for h in HEADERS}
    for idx, cell in enumerate(row):
        n = norm(cell)
        # Alias handling first
        if n in alias_map and alias_map[n] not in mapping.values():
            mapping[idx] = alias_map[n]
            continue

        # Direct match to master headers using normalized form
        if n in normalized_master and normalized_master[n] not in mapping.values():
            mapping[idx] = normalized_master[n]
            continue
    return mapping


def pad(row: List[str], size: int) -> List[str]:
    row = row[:size]
    if len(row) < size:
        row += [''] * (size - len(row))
    return row


def process_csv(path: Path, collector: List[Dict[str, str]]):
    current_header_map: Optional[Dict[int, str]] = None
    current_manufacturer = ''
    current_series = ''

    def empty_record() -> Dict[str, str]:
        return {h: '' for h in HEADERS}

    # Determine global attributes based on filename
    fname = path.stem.lower()

    file_anamorphic = 'Anamorphic' if 'anamorphic' in fname else 'Spherical'

    def detect_format(name: str) -> str:
        if '65mm' in name:
            return '65mm'
        if 'fullframe' in name or 'full_frame' in name or 'full frame' in name:
            return 'Full Frame'
        if 'super35' in name or 'super_35' in name or 'super 35' in name:
            return 'Super 35'
        if '16mm' in name:
            return '16mm Format'
        return ''

    file_format = detect_format(fname)

    with path.open(newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        for raw in reader:
            if not any(raw):
                continue

            if looks_like_section_header(raw):
                current_header_map = build_header_map(raw)
                # Determine manufacturer/series column indices for propagation
                man_idx = next((idx for idx, hdr in current_header_map.items() if hdr == 'Manufacturer'), None)
                ser_idx = next((idx for idx, hdr in current_header_map.items() if hdr == 'Series'), None)
                continue

            if current_header_map is None:
                continue  # skip lines before first header

            record = empty_record()
            # Map section-specific headers
            for idx, cell in enumerate(raw):
                if idx in current_header_map:
                    record[current_header_map[idx]] = cell.strip()

            # Column 0 (Manufacturer) & 1 (Series) are fixed per spec
            manufacturer_cell = raw[0].strip() if len(raw) > 0 else ''
            series_cell = raw[1].strip() if len(raw) > 1 else ''

            if manufacturer_cell:
                current_manufacturer = manufacturer_cell
            record['Manufacturer'] = current_manufacturer

            if series_cell:
                current_series = series_cell
            record['Series'] = current_series

            # Macro handling in focal length
            focal_str = record['Focal Length']
            if focal_str and 'macro' in focal_str.lower():
                cleaned = re.sub(r'(?i)\bmacro\b', '', focal_str).strip()
                record['Focal Length'] = cleaned
                if 'Macro' not in record['Notes']:
                    record['Notes'] = (record['Notes'] + '; ' if record['Notes'] else '') + 'Macro'
                focal_str = cleaned

            # Prime / Zoom detection
            if focal_str:
                record['Prime / Zoom / Special'] = 'Zoom' if '-' in focal_str else 'Prime'

            # Clean not in inventory
            if record['Notes']:
                record['Notes'] = re.sub(r'(?i)\bnot\s+(?:currently\s+)?in\s+inventory\b', '', record['Notes']).strip().strip(',;')

            # Apply file-level defaults if fields are empty
            if not record['Anamorphic / Spherical']:
                record['Anamorphic / Spherical'] = file_anamorphic

            if not record['Format'] and file_format:
                record['Format'] = file_format

            collector.append(record)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main(folder: Path):
    if not folder.exists() or not folder.is_dir():
        sys.exit(f"Folder not found: {folder}")

    all_rows: List[Dict[str, str]] = []
    csv_files = sorted(folder.glob('*.csv'))
    if not csv_files:
        sys.exit(f"No CSV files found in {folder}")

    for csv_path in csv_files:
        process_csv(csv_path, all_rows)

    with OUTPUT_FILE.open('w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=HEADERS)
        writer.writeheader()
        writer.writerows(all_rows)

    print(f"Flattened {len(all_rows)} rows from {len(csv_files)} CSV(s) → {OUTPUT_FILE}")

if __name__ == '__main__':
    target_folder = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_FOLDER
    main(target_folder) 