#!/usr/bin/env python3
"""
ESC Raw Lense Parse
-------------------
Explode the one-column "ESC Raw Lenses.csv" list into the master 36-column
format used by Flattened_Lens_Inventory.csv.

Usage:
    python3 esc_raw_lense_parse.py  [input_csv] [output_csv]
If paths are omitted it defaults to the file next to the script:
    input : ESC Raw Lenses.csv
    output: ESC_Raw_Lenses_Flat.csv
"""
import csv
import re
import sys
from pathlib import Path
from typing import Dict, List

PROJECT_DIR = Path(__file__).parent
INPUT_DEFAULT = PROJECT_DIR / "ESC Raw Lenses.csv"
OUTPUT_DEFAULT = PROJECT_DIR / "ESC_Raw_Lenses_Flat.csv"

HEADERS: List[str] = [
    'Manufacturer', 'Series', 'Focal Length', 'T-Stop', 'Prime / Zoom / Special',
    'Format', 'Mount', 'Anamorphic / Spherical', 'Anamorphic Squeeze Factor',
    'Anamorphic Location', 'Housing', 'Front Diameter (mm)', 'Close Focus',
    'Length (in)', 'Film Compatibility', 'Image Circle (mm)', 'Iris Blade Count',
    'Extender', 'LDS', 'i/Data', 'Support Recommended', 'Support Post Length (mm)',
    'Weight (lbs)', 'Manufacture Year', 'Expander', 'Heden Motor Size', 'Size',
    'Notes', 'Look', 'Use Case', 'Bokeh', 'Flare', 'Focus Falloff', 'Breathing',
    'Focus Scale', 'Original Name'
]

SERIES_KEYWORDS = [
    'Compact Prime CP3', 'Compact Prime CP2', 'Compact Prime',
    'Master Prime', 'Ultra Prime', 'Super Speed', 'Standard Speed',
    'S4', 'S4/i', 'S5', 'S7', 'S8', 'Panchro', 'Speed Panchro', 'Varotal',
    'Summilux C', 'Summicron C', 'Summilux-C', 'Summicron-C', 'Cine'
]

HOUSING_MANUFACTURERS = [
    'TLS', 'Ancient Optics', 'GL Optics', 'Optex', 'Whitepoint Optics', 
    'Zero Optik', 'Works Cameras', 'Cinescope'
]

FORMAT_KEYWORDS = [
    'FF', '16mm Format', 'S16', 'S35', 'Super 35', 'Full Frame', '65mm Format', 'Large Format',
    'Super 16', 'Super 35mm'
]

MOUNT_KEYWORDS = [
    'LPL Mount', 'XPL Mount', 'PL Mount', 'Canon EF', 'Sony E', 'Nikon Z',
    'LPL', 'XPL', 'E mount', 'EF Mount', 'Z mount', 'EF', 'E-Mount', 'Z-Mount', 'PL'
]

ANAMORPHIC_KEYWORDS = [
    'Anamorphic', 'Spherical'
]

# Manufacturer-Series Dictionary for filling in blanks
MANUFACTURER_SERIES_DICT = {
    'Zeiss': ['Master Prime', 'Ultra Prime', 'CP2', 'Supreme Prime', 'CP3', 'Compact Prime CP2', 'Compact Prime CP3', 'Standard Speed', 'Super Speed'],
    'Canon': ['CN-E Cinema Primes', 'K35', 'K-35'],
    'Nikon': ['Nikkor Z', 'AI-S Cine-Mod', 'Nikkor'],
    'Leica': ['Thalia', 'R'],
    'Leitz': ['Summilux C', 'Summicron C', 'Leitz Prime', 'Hugo', 'R'],
    'Fujinon': ['MK Series', 'Premista', 'Premier'],
    'Angenieux': ['Optimo', 'EZ Series', 'HR', 'A-2S'],
    'Cooke': ['S4/i', 'Anamorphic/i', 'S4', 'S5', 'S7', 'S8', 'Panchro', 'Speed Panchro', 'Varotal'],
    'Schneider': ['Xenon FF', 'Cine-Xenar III'],
    'Tokina': ['Vista Primes', 'ATX Cinema'],
    'Sigma': ['Cine FF High Speed', 'Art Series'],
    'Tamron': ['SP Series', 'Di Series'],
    'Rokinon': ['XEEN', 'Cine DS'],
    'Samyang': ['VDSLR', 'XEEN'],
    'Irix': ['Cine Series', 'Dragonfly'],
    'Venus': ['Laowa Cine', 'Zero-D'],
    'Mitakon': ['Speedmaster', 'Creator'],
    'Meike': ['Cinema Prime', 'Classic Cine'],
    '7Artisans': ['Vision Series', 'Photoelectric'],
    'Laowa': ['Zero-D Cine', 'Probe Lens', 'FF Ranger'],
    'Voigtländer': ['Nokton', 'APO-Lanthar', 'Heliar', 'Ultra Wide Heliar', 'Super Wide Heliar'],
    'Voigtlander': ['Nokton', 'APO-Lanthar', 'Heliar', 'Ultra Wide Heliar', 'Super Wide Heliar'],
    'Hawk': ['V-Lite', 'V-Plus Anamorphic'],
    'Masterbuilt': ['Masterbuilt Primes', 'Legacy Anamorphic'],
    'Caldwell': ['Chameleon Anamorphic', 'IBÉ Optics-Caldwell 1.79x', 'Chameleon'],
    'Xelmus': ['Apollo Anamorphic', 'Helium'],
    'DZOFilms': ['VESPID Prime', 'Pictor Zoom'],
    'Atlas': ['Orion Series', 'Mercury Series'],
    'Tribe7': ['Blackwing7', 'T-Tuned Primes'],
    'Kowa': ['Anamorphic Prominar', 'Cine Prominar', 'Cine'],
    'Gecko-Cam': ['Genesis G35'],
    'Sony': ['G Master', 'CineAlta'],
    'Panavision': ['Primo', 'Ultra Panatar'],
    'Vantage': ['Hawk V-Lite', 'Vantage One T1'],
    'Iscorama': ['Iscorama 36', 'Iscorama 54'],
    'Century': ['Century Optics Anamorphic', 'Pro Series'],
    'Statera': ['Statera Primes', 'Statera Macro'],
    'Master': ['Master Primes', 'Master Anamorphic', 'Master Prime'],
    'Lomo': ['LOMO Roundfront', 'LOMO Squarefront'],
    'ARRI': ['Signature Prime', 'Master Prime', 'Ultra Prime'],
    'Arri': ['Signature Prime', 'Master Prime', 'Ultra Prime']
}

EXTRA_FLAGS = {
    'lds': 'LDS',
    'macro': 'Macro',
    'mfx': 'MFX',
    'uncoated': 'Uncoated',
    'special fx': 'Special FX',
}

def empty_row() -> Dict[str, str]:
    return {h: '' for h in HEADERS}


def clean_text(text: str) -> str:
    return re.sub(r'[\s\-_,;]+', ' ', text).strip()


def detect_housing(original: str, manufacturer: str) -> str:
    """
    Detect housing manufacturer from the lens name.
    If a housing manufacturer is found, return it regardless of lens manufacturer.
    Otherwise return "Original Housing".
    """
    for housing_mfg in HOUSING_MANUFACTURERS:
        if re.search(re.escape(housing_mfg), original, re.I):
            return housing_mfg
    
    return "Original Housing"


def fill_blanks_with_dict(row: Dict[str, str], original: str) -> Dict[str, str]:
    """
    Fill in blank manufacturer or series using the manufacturer-series dictionary.
    Only fills blanks, never overwrites existing data.
    """
    # If both manufacturer and series are filled, no need to do anything
    if row['Manufacturer'] and row['Series']:
        return row
    
    # If we have a manufacturer but no series, try to find a matching series
    if row['Manufacturer'] and not row['Series']:
        manufacturer = row['Manufacturer'].strip()
        if manufacturer in MANUFACTURER_SERIES_DICT:
            # Look for any of the known series in the original name
            for series in MANUFACTURER_SERIES_DICT[manufacturer]:
                if re.search(re.escape(series), original, re.I):
                    row['Series'] = series
                    break
    
    # If we have a series but no manufacturer, try to find a matching manufacturer
    elif row['Series'] and not row['Manufacturer']:
        series = row['Series'].strip()
        for manufacturer, series_list in MANUFACTURER_SERIES_DICT.items():
            if series in series_list:
                row['Manufacturer'] = manufacturer
                break
    
    # If both are blank, try to infer from the original name
    elif not row['Manufacturer'] and not row['Series']:
        # Look for any manufacturer in the original name
        for manufacturer, series_list in MANUFACTURER_SERIES_DICT.items():
            if re.search(re.escape(manufacturer), original, re.I):
                row['Manufacturer'] = manufacturer
                # Try to find a matching series
                for series in series_list:
                    if re.search(re.escape(series), original, re.I):
                        row['Series'] = series
                        break
                break
    
    return row


def parse_line(line: str) -> Dict[str, str]:
    row = empty_row()
    original = line.strip()
    if not original:
        return row
    
    # Store the original unprocessed name in column AJ (Original Name)
    row['Original Name'] = original

    # PRIORITY 1: focal length detection (including zoom ranges)
    # Check for zoom range first (e.g., 25-100mm, 25mm-100mm, 25-100)
    zoom_match = re.search(r'(\d+(?:\.\d+)?)\s*[-–]\s*(\d+(?:\.\d+)?)\s*mm', original, re.I)
    if zoom_match:
        row['Focal Length'] = f"{zoom_match.group(1)}mm-{zoom_match.group(2)}mm"
    else:
        # Check for single focal length
        fl_match = re.search(r'(\d+(?:\.\d+)?)\s*mm', original, re.I)
        if not fl_match:
            fl_match = re.match(r'(\d+(?:\.\d+)?)', original)
        if fl_match:
            row['Focal Length'] = f"{fl_match.group(1)}mm"

    # PRIORITY 2: T-Stop
    t_match = re.search(r'T\s*(\d+(?:\.\d+)?)', original, re.I)
    if t_match:
        row['T-Stop'] = t_match.group(1)



    row['Prime / Zoom / Special'] = 'Prime'

    # PRIORITY 3: Format detection (avoiding conflicts with focal length)
    for format_key in FORMAT_KEYWORDS:
        if re.search(re.escape(format_key), original, re.I):
            # Double-check this isn't actually a focal length
            # If it's a simple number+mm pattern and we already have a focal length, skip it
            if re.match(r'^\d+mm$', format_key, re.I) and row['Focal Length']:
                continue
            row['Format'] = format_key
            break

    # PRIORITY 4: Manufacturer detection - typically comes after focal length
    if row['Focal Length']:
        # Split after focal length and look for manufacturer
        after_fl = original.split(row['Focal Length'], 1)[1] if row['Focal Length'] in original else original
        after_fl = re.sub(r'^\s*mm\s*', '', after_fl, flags=re.I)
        
        # Look for common manufacturer names
        manufacturer_patterns = [
            r'\b(Zeiss|Canon|Nikon|Leica|Leitz|Fujinon|Angenieux|Cooke|Schneider|Tokina|Sigma|Tamron|Rokinon|Samyang|Irix|Venus|Mitakon|Meike|7Artisans|Laowa|Voigtländer|Voigtlander|Hawk|Masterbuilt|Caldwell|Xelmus|DZOFilms|Atlas|Tribe7|Kowa|Gecko-Cam|Sony|Panavision|Vantage|Iscorama|Century|Statera|Master|Lomo|V|ARRI|Arri)\b',
            r'\b(ARRI|Arri|Arri)\b',
            r'\b(Zeiss)\b'
        ]
        
        for pattern in manufacturer_patterns:
            mfg_match = re.search(pattern, after_fl, re.I)
            if mfg_match:
                row['Manufacturer'] = mfg_match.group(1)
                break

    # PRIORITY 5: Series detection
    # Special handling for Cooke series - look for S4/i pattern specifically
    if row['Manufacturer'] and 'Cooke' in row['Manufacturer']:
        if re.search(r'S4/i', original, re.I):
            row['Series'] = 'S4/i'
        elif re.search(r'S4', original, re.I):
            row['Series'] = 'S4'
        elif re.search(r'5i', original, re.I):
            row['Series'] = '5i'
    
    # Special handling for Leitz/Leica series
    if row['Manufacturer'] and ('Leitz' in row['Manufacturer'] or 'Leica' in row['Manufacturer']):
        if re.search(r'Summilux-C|Summilux C', original, re.I):
            row['Series'] = 'Summilux C'
        elif re.search(r'Summicron-C|Summicron C', original, re.I):
            row['Series'] = 'Summicron C'
        elif re.search(r'Cine', original, re.I):
            row['Series'] = 'Cine'
        elif re.search(r'Leica-R|\bR\b', original, re.I):
            row['Series'] = 'R'
    
    # General series detection
    if not row['Series']:
        for key in SERIES_KEYWORDS:
            if re.search(re.escape(key), original, re.I):
                row['Series'] = key
                break
    
    # Fallback: extract series from the beginning of the name after focal length
    if not row['Series']:
        if row['Focal Length']:
            # Split after focal length and take the first word/phrase
            after_fl = original.split(row['Focal Length'], 1)[1] if row['Focal Length'] in original else original
            after_fl = re.sub(r'^\s*mm\s*', '', after_fl, flags=re.I).strip()
            
            # Split by common separators and take the first meaningful word
            parts = re.split(r'[\s\-_,;()]+', after_fl)
            series_candidate = parts[0].strip() if parts else ""
            
            # Clean up the series candidate
            if series_candidate:
                # Remove manufacturer if found
                if row['Manufacturer']:
                    series_candidate = re.sub(re.escape(row['Manufacturer']), '', series_candidate, flags=re.I)
                # Remove format keywords
                for format_key in FORMAT_KEYWORDS:
                    series_candidate = re.sub(re.escape(format_key), '', series_candidate, flags=re.I)
                # Remove anamorphic/spherical keywords
                for anamorphic_key in ANAMORPHIC_KEYWORDS:
                    series_candidate = re.sub(re.escape(anamorphic_key), '', series_candidate, flags=re.I)
                # Remove anamorphic squeeze factors
                series_candidate = re.sub(r'\d+(?:\.\d+)?\s*x', '', series_candidate, flags=re.I)
                # Remove mount information
                for mount_key in MOUNT_KEYWORDS:
                    mount_pattern = r'\(?\s*' + re.escape(mount_key) + r'\s*\)?'
                    series_candidate = re.sub(mount_pattern, '', series_candidate, flags=re.I)
                # Remove housing manufacturers
                for housing_mfg in HOUSING_MANUFACTURERS:
                    series_candidate = re.sub(re.escape(housing_mfg), '', series_candidate, flags=re.I)
                
                series_candidate = series_candidate.strip()
                if series_candidate:
                    # Skip if the candidate is just a number (likely part of focal length)
                    if re.match(r'^\d+(?:\.\d+)?$', series_candidate):
                        pass
                    # Filter out single-letter series candidates unless they're for Leitz/Leica
                    elif len(series_candidate) == 1 and not (row['Manufacturer'] and ('Leitz' in row['Manufacturer'] or 'Leica' in row['Manufacturer'])):
                        # Skip single-letter series for non-Leitz/Leica lenses
                        pass
                    else:
                        row['Series'] = series_candidate

    # Determine Prime/Zoom/Special based on focal length and keywords
    if row['Focal Length']:
        # Check if focal length contains a range (zoom lens)
        if re.search(r'\d+mm?\s*[-–]\s*\d+mm?', row['Focal Length'], re.I):
            row['Prime / Zoom / Special'] = 'Zoom'
        else:
            row['Prime / Zoom / Special'] = 'Prime'
    else:
        row['Prime / Zoom / Special'] = 'Prime'
    
    # Override with keyword detection - if "Zoom" is in the name, it's a zoom lens
    if re.search(r'\bzoom\b', original, re.I):
        row['Prime / Zoom / Special'] = 'Zoom'

    # PRIORITY 6: Mount detection (including parentheses)
    for mount_key in MOUNT_KEYWORDS:
        # Look for mount in parentheses or as standalone
        mount_pattern = r'\(?\s*' + re.escape(mount_key) + r'\s*\)?'
        if re.search(mount_pattern, original, re.I):
            row['Mount'] = mount_key
            break
    # Default to PL if no mount found
    if not row['Mount']:
        row['Mount'] = 'PL'

    # PRIORITY 7: Anamorphic/Spherical detection and squeeze factor
    for anamorphic_key in ANAMORPHIC_KEYWORDS:
        if re.search(re.escape(anamorphic_key), original, re.I):
            row['Anamorphic / Spherical'] = anamorphic_key
            break
    
    # If anamorphic, look for squeeze factor (1.8x, 2x, 1.5x, etc.)
    if row['Anamorphic / Spherical'] == 'Anamorphic':
        squeeze_match = re.search(r'(\d+(?:\.\d+)?)\s*x', original, re.I)
        if squeeze_match:
            row['Anamorphic Squeeze Factor'] = f"{squeeze_match.group(1)}x"

    # PRIORITY 8: Housing detection
    row['Housing'] = detect_housing(original, row['Manufacturer'])

    # PRIORITY 9: Special exceptions and additional detection
    # Handle "Kooky Cooke" exception - if "Kooky Cooke" appears, add it to notes
    if re.search(r'Kooky Cooke', original, re.I):
        if not row['Notes']:
            row['Notes'] = 'Kooky Cooke'
        else:
            row['Notes'] = row['Notes'] + '; Kooky Cooke'
    
    # Detect i/Data for Cooke lenses
    if row['Manufacturer'] and 'Cooke' in row['Manufacturer'] and re.search(r'i/Data|iData', original, re.I):
        row['i/Data'] = 'Yes'



    # Flags
    notes = []
    for token, label in EXTRA_FLAGS.items():
        if token in original.lower():
            if label == 'LDS':
                row['LDS'] = 'Yes'
            else:
                notes.append(label)
    if notes:
        row['Notes'] = ', '.join(notes)

    # ---------------------------------------------------------------
    # Capture any leftover descriptive text into notes
    # ---------------------------------------------------------------
    residual = original
    
    # Remove all detected components systematically
    components_to_remove = []
    
    # Add manufacturer
    if row['Manufacturer']:
        components_to_remove.append(row['Manufacturer'])
    
    # Add series
    if row['Series']:
        components_to_remove.append(row['Series'])
    
    # Add focal length
    if row['Focal Length']:
        components_to_remove.append(row['Focal Length'])
    
    # Add T-stop
    if row['T-Stop']:
        components_to_remove.append(f"T{row['T-Stop']}")
    
    # Add anamorphic squeeze factor
    if row['Anamorphic Squeeze Factor']:
        components_to_remove.append(row['Anamorphic Squeeze Factor'])
    
    # Add format keywords (but preserve focal length patterns)
    for format_key in FORMAT_KEYWORDS:
        if not re.match(r'^\d+mm$', format_key, re.I):  # Skip simple focal length patterns
            components_to_remove.append(format_key)
    
    # Add mount keywords
    for mount_key in MOUNT_KEYWORDS:
        components_to_remove.append(mount_key)
    
    # Add anamorphic/spherical keywords
    for anamorphic_key in ANAMORPHIC_KEYWORDS:
        components_to_remove.append(anamorphic_key)
    
    # Add housing manufacturers
    for housing_mfg in HOUSING_MANUFACTURERS:
        components_to_remove.append(housing_mfg)
    
    # Add flag tokens
    for token in EXTRA_FLAGS.keys():
        components_to_remove.append(token)
    
    # Remove all components from residual text
    for component in components_to_remove:
        # Use word boundaries to avoid partial matches
        pattern = r'\b' + re.escape(component) + r'\b'
        residual = re.sub(pattern, '', residual, flags=re.I)
    
    # Remove mount patterns with parentheses
    for mount_key in MOUNT_KEYWORDS:
        mount_pattern = r'\(?\s*' + re.escape(mount_key) + r'\s*\)?'
        residual = re.sub(mount_pattern, '', residual, flags=re.I)
    
    # Remove anamorphic squeeze factors
    residual = re.sub(r'\d+(?:\.\d+)?\s*x', '', residual, flags=re.I)
    
    # Remove i/Data patterns
    residual = re.sub(r'i/Data|iData', '', residual, flags=re.I)
    
    # Remove Leitz/Leica series patterns
    residual = re.sub(r'Summilux-C|Summilux C', '', residual, flags=re.I)
    residual = re.sub(r'Summicron-C|Summicron C', '', residual, flags=re.I)
    residual = re.sub(r'Cine', '', residual, flags=re.I)
    # Only remove "R" if it's a Leitz/Leica lens
    if row['Manufacturer'] and ('Leitz' in row['Manufacturer'] or 'Leica' in row['Manufacturer']):
        residual = re.sub(r'Leica-R|R', '', residual, flags=re.I)
    
    # Remove the detected focal length from residual text
    if row['Focal Length']:
        # Remove the exact focal length that was detected
        residual = re.sub(re.escape(row['Focal Length']), '', residual, flags=re.I)
        
        # For zoom lenses, also remove the individual focal lengths
        if '-' in row['Focal Length']:
            # Extract the two focal lengths from the range
            parts = row['Focal Length'].split('-')
            if len(parts) == 2:
                first_fl = parts[0].replace('mm', '')
                second_fl = parts[1].replace('mm', '')
                # Remove both individual focal lengths
                residual = re.sub(r'\b' + re.escape(first_fl) + r'\b', '', residual, flags=re.I)
                residual = re.sub(r'\b' + re.escape(second_fl) + r'\b', '', residual, flags=re.I)
                # Also remove with mm suffix
                residual = re.sub(r'\b' + re.escape(first_fl) + r'mm\b', '', residual, flags=re.I)
                residual = re.sub(r'\b' + re.escape(second_fl) + r'mm\b', '', residual, flags=re.I)
        
        # Also remove the focal length without "mm" suffix if it was added
        focal_without_mm = row['Focal Length'].replace('mm', '')
        if focal_without_mm != row['Focal Length']:
            residual = re.sub(re.escape(focal_without_mm), '', residual, flags=re.I)
    
    # Clean up common separators and formatting
    residual = re.sub(r'[\s\-_,;]+', ' ', residual)  # Replace separators with single space
    residual = residual.replace('(', '').replace(')', '')
    residual = residual.replace('[', '').replace(']', '')
    residual = residual.replace('+', ' ')
    residual = clean_text(residual)

    # Add any remaining text to notes
    if residual:
        row['Notes'] = (row['Notes'] + '; ' if row['Notes'] else '') + residual

    # Fill in blanks using manufacturer-series dictionary as a last resort
    row = fill_blanks_with_dict(row, original)

    return row


def main(input_path: Path, output_path: Path):
    # Check if output file already exists and delete it
    if output_path.exists():
        output_path.unlink()
        print(f"Deleted existing file: {output_path}")
    
    rows: List[Dict[str, str]] = []
    with input_path.open(encoding='utf-8') as f:
        for line in f:
            parsed = parse_line(line)
            # Skip completely blank rows
            if any(parsed.values()):
                rows.append(parsed)

    with output_path.open('w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=HEADERS)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Parsed {len(rows)} raw lenses → {output_path}")


if __name__ == '__main__':
    inp = Path(sys.argv[1]) if len(sys.argv) > 1 else INPUT_DEFAULT
    outp = Path(sys.argv[2]) if len(sys.argv) > 2 else OUTPUT_DEFAULT
    main(inp, outp) 