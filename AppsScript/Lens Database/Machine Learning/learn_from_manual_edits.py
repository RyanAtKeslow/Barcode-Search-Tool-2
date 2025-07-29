#!/usr/bin/env python3
"""
Learn From Manual Edits
=======================

This script reads the Manual Edits file and learns new parsing patterns
from your manual corrections to improve the parser.
"""

import pandas as pd
import json
from pathlib import Path
from typing import Dict, List, Set
import re

def load_manual_edits(file_path: str) -> pd.DataFrame:
    """Load the manual edits file"""
    try:
        df = pd.read_csv(file_path)
        print(f"Loaded {len(df)} manual edits from {file_path}")
        return df
    except Exception as e:
        print(f"Error loading manual edits: {e}")
        return pd.DataFrame()

def analyze_manual_edits(df: pd.DataFrame) -> Dict:
    """Analyze manual edits to learn new patterns"""
    patterns = {
        'manufacturers': {},
        'series': {},
        'mounts': {},
        'formats': {},
        't_stops': {},
        'focal_lengths': {},
        'lens_types': {},
        'missed_detections': {
            'manufacturers': [],
            'series': []
        }
    }
    
    if df.empty:
        return patterns
    
    # Analyze each field for patterns
    for _, row in df.iterrows():
        original_name = str(row.get('Original Name', ''))
        manufacturer = str(row.get('Manufacturer', ''))
        series = str(row.get('Series', ''))
        mount = str(row.get('Mount', ''))
        format_info = str(row.get('Format', ''))
        t_stop = str(row.get('T-Stop', ''))
        focal_length = str(row.get('Focal Length', ''))
        lens_type = str(row.get('Prime / Zoom / Special', ''))
        

        
        # Learn manufacturer patterns
        if manufacturer and manufacturer.lower() not in ['nan', '']:
            patterns['manufacturers'][manufacturer.lower()] = patterns['manufacturers'].get(manufacturer.lower(), 0) + 1
        
        # Skip single-letter 'r' as series
        if series and series.lower() not in ['nan', ''] and series.strip().lower() != 'r':
            patterns['series'][series.lower()] = patterns['series'].get(series.lower(), 0) + 1
        
        # Learn mount patterns
        if mount and mount.lower() not in ['nan', '']:
            patterns['mounts'][mount.lower()] = patterns['mounts'].get(mount.lower(), 0) + 1
        
        # Learn format patterns
        if format_info and format_info.lower() not in ['nan', '']:
            patterns['formats'][format_info.lower()] = patterns['formats'].get(format_info.lower(), 0) + 1
        
        # Learn T-stop patterns
        if t_stop and t_stop.lower() not in ['nan', '']:
            patterns['t_stops'][t_stop.lower()] = patterns['t_stops'].get(t_stop.lower(), 0) + 1
        
        # Learn focal length patterns
        if focal_length and focal_length.lower() not in ['nan', '']:
            patterns['focal_lengths'][focal_length.lower()] = patterns['focal_lengths'].get(focal_length.lower(), 0) + 1
        
        # Learn lens type patterns
        if lens_type and lens_type.lower() not in ['nan', '']:
            patterns['lens_types'][lens_type.lower()] = patterns['lens_types'].get(lens_type.lower(), 0) + 1
        
        # Check for missed detections
        original_lower = original_name.lower()
        if manufacturer and manufacturer.lower() not in ['nan', '']:
            if manufacturer.lower() not in original_lower:
                patterns['missed_detections']['manufacturers'].append({
                    'original': original_name,
                    'expected_manufacturer': manufacturer,
                    'series': series
                })
        
        if series and series.lower() not in ['nan', '']:
            if series.lower() not in original_lower:
                patterns['missed_detections']['series'].append({
                    'original': original_name,
                    'expected_series': series,
                    'manufacturer': manufacturer
                })
    
    return patterns

def extract_patterns_from_names(df: pd.DataFrame) -> Dict:
    """Extract patterns from original names that correspond to manual corrections"""
    patterns = {
        'manufacturer_patterns': {},
        'series_patterns': {},
        'mount_patterns': {},
        'format_patterns': {}
    }
    
    if df.empty:
        return patterns
    
    for _, row in df.iterrows():
        original_name = str(row.get('Original Name', '')).lower()
        manufacturer = str(row.get('Manufacturer', '')).lower()
        series = str(row.get('Series', '')).lower()
        mount = str(row.get('Mount', '')).lower()
        format_info = str(row.get('Format', '')).lower()
        
        # Find patterns in original name that correspond to manual corrections
        if manufacturer and manufacturer not in ['nan', '']:
            # Look for manufacturer patterns in original name
            for word in original_name.split():
                if word in manufacturer or manufacturer in word:
                    patterns['manufacturer_patterns'][word] = patterns['manufacturer_patterns'].get(word, 0) + 1
        
        if series and series not in ['nan', ''] and series.strip() != 'r':
            # Look for series patterns in original name
            for word in original_name.split():
                if word in series or series in word:
                    patterns['series_patterns'][series] = patterns['series_patterns'].get(series, 0) + 1
        
        if mount and mount not in ['nan', '']:
            # Look for mount patterns in original name
            mount_patterns = re.findall(r'\([^)]*\)', original_name)  # Find parentheses patterns
            for pattern in mount_patterns:
                if mount in pattern.lower():
                    patterns['mount_patterns'][pattern] = patterns['mount_patterns'].get(pattern, 0) + 1
        
        if format_info and format_info not in ['nan', '']:
            # Look for format patterns in original name
            for word in original_name.split():
                if word in format_info or format_info in word:
                    patterns['format_patterns'][word] = patterns['format_patterns'].get(word, 0) + 1
    
    return patterns

def generate_improved_parser(manual_patterns: Dict, name_patterns: Dict) -> str:
    """Generate improved parser code based on learned patterns"""
    
    # Generate new manufacturer patterns
    new_manufacturers = []
    for manufacturer, count in sorted(manual_patterns['manufacturers'].items(), key=lambda x: x[1], reverse=True):
        if count >= 2:  # Only include patterns that appear multiple times
            new_manufacturers.append(f"    '{manufacturer}': ['{manufacturer}'],")
    
    # Generate new series patterns
    new_series = []
    for series, count in sorted(manual_patterns['series'].items(), key=lambda x: x[1], reverse=True):
        if count >= 2:  # Only include patterns that appear multiple times
            new_series.append(f"    '{series}': ['{series}'],")
    
    # Generate new mount patterns
    new_mounts = []
    for mount, count in sorted(manual_patterns['mounts'].items(), key=lambda x: x[1], reverse=True):
        if count >= 2:  # Only include patterns that appear multiple times
            new_mounts.append(f"    '{mount}': ['{mount}'],")
    
    # Generate new format patterns
    new_formats = []
    for format_info, count in sorted(manual_patterns['formats'].items(), key=lambda x: x[1], reverse=True):
        if count >= 2:  # Only include patterns that appear multiple times
            new_formats.append(f"    '{format_info}': ['{format_info}'],")
    
    # Generate the improved parser code
    improved_code = f"""
# Improved parser patterns learned from manual edits
# Generated automatically from Manual Edits file

# New manufacturer patterns:
{chr(10).join(new_manufacturers)}

# New series patterns:
{chr(10).join(new_series)}

# New mount patterns:
{chr(10).join(new_mounts)}

# New format patterns:
{chr(10).join(new_formats)}

# Name-based patterns found:
"""
    
    # Add name-based patterns
    for pattern_type, patterns in name_patterns.items():
        improved_code += f"\n# {pattern_type}:\n"
        for pattern, count in sorted(patterns.items(), key=lambda x: x[1], reverse=True):
            if count >= 2:  # Only include patterns that appear multiple times
                improved_code += f"#   '{pattern}': {count} occurrences\n"
    
    return improved_code

def main():
    """Main function to learn from manual edits"""
    print("Learning From Manual Edits")
    print("=" * 50)
    
    # Load manual edits file
    manual_edits_file = Path("Manual Edits.csv")
    
    if not manual_edits_file.exists():
        print(f"Manual Edits file not found: {manual_edits_file}")
        print("Please ensure the file exists in the current directory.")
        return
    
    # Load and analyze manual edits
    df = load_manual_edits(str(manual_edits_file))
    if df.empty:
        return
    
    # Analyze patterns
    print("Analyzing manual corrections...")
    manual_patterns = analyze_manual_edits(df)
    name_patterns = extract_patterns_from_names(df)
    
    # Generate improved parser
    print("Generating improved parser patterns...")
    improved_code = generate_improved_parser(manual_patterns, name_patterns)
    
    # Save the improved patterns
    output_file = "learned_patterns.py"
    with open(output_file, 'w') as f:
        f.write(improved_code)
    
    # Save patterns as JSON for reference
    json_file = "learned_patterns.json"
    with open(json_file, 'w') as f:
        json.dump({
            'manual_patterns': manual_patterns,
            'name_patterns': name_patterns
        }, f, indent=2)
    
    print(f"\nAnalysis complete!")
    print(f"Learned patterns saved to: {output_file}")
    print(f"Pattern data saved to: {json_file}")
    
    # Show summary
    print(f"\n=== PATTERN SUMMARY ===")
    print(f"New manufacturers found: {len([k for k, v in manual_patterns['manufacturers'].items() if v >= 2])}")
    print(f"New series found: {len([k for k, v in manual_patterns['series'].items() if v >= 2])}")
    print(f"New mounts found: {len([k for k, v in manual_patterns['mounts'].items() if v >= 2])}")
    print(f"New formats found: {len([k for k, v in manual_patterns['formats'].items() if v >= 2])}")
    
    # Show top patterns
    print(f"\n=== TOP MANUFACTURERS ===")
    for manufacturer, count in sorted(manual_patterns['manufacturers'].items(), key=lambda x: x[1], reverse=True)[:10]:
        if count >= 2:
            print(f"  {manufacturer}: {count}")
    
    print(f"\n=== TOP SERIES ===")
    for series, count in sorted(manual_patterns['series'].items(), key=lambda x: x[1], reverse=True)[:10]:
        if count >= 2:
            print(f"  {series}: {count}")
    
    print(f"\n=== TOP MOUNTS ===")
    for mount, count in sorted(manual_patterns['mounts'].items(), key=lambda x: x[1], reverse=True)[:10]:
        if count >= 2:
            print(f"  {mount}: {count}")
    
    print(f"\n=== TOP FORMATS ===")
    for format_info, count in sorted(manual_patterns['formats'].items(), key=lambda x: x[1], reverse=True)[:10]:
        if count >= 2:
            print(f"  {format_info}: {count}")
    
    # Show missed detections
    print(f"\n=== MISSED MANUFACTURER DETECTIONS ===")
    for missed in manual_patterns['missed_detections']['manufacturers'][:20]:  # Show first 20
        print(f"  Original: '{missed['original']}'")
        print(f"    Expected: {missed['expected_manufacturer']} (Series: {missed['series']})")
        print()
    
    print(f"\n=== MISSED SERIES DETECTIONS ===")
    for missed in manual_patterns['missed_detections']['series'][:20]:  # Show first 20
        print(f"  Original: '{missed['original']}'")
        print(f"    Expected: {missed['expected_series']} (Manufacturer: {missed['manufacturer']})")
        print()

if __name__ == "__main__":
    main() 