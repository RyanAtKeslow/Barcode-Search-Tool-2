#!/usr/bin/env python3
"""
Lens Review and Improvement Tool
================================

This script helps you review lenses that need manual review and improve
the parser's confidence scores by adding custom patterns.
"""

import pandas as pd
import json
from pathlib import Path
from simple_lens_parser import SimpleLensParser

def load_parsed_data():
    """Load the parsed lens data"""
    try:
        df = pd.read_csv("parsed_lenses_output.csv")
        print(f"Loaded {len(df)} parsed lenses")
        return df
    except FileNotFoundError:
        print("Error: parsed_lenses_output.csv not found")
        print("Please run process_existing_data.py first")
        return None

def show_lenses_needing_review(df, limit=10):
    """Show lenses that need review"""
    needs_review = df[df['Needs Review'] == True].copy()
    
    if len(needs_review) == 0:
        print("No lenses need review!")
        return needs_review
    
    print(f"\n=== LENSES NEEDING REVIEW ({len(needs_review)} total) ===")
    print("Showing first", limit, "lenses:")
    print()
    
    for idx, row in needs_review.head(limit).iterrows():
        print(f"{idx+1}. {row['Original Name']}")
        print(f"   Confidence: {row['Confidence Score']:.3f}")
        print(f"   Manufacturer: {row['Manufacturer']}")
        print(f"   Series: {row['Series']}")
        print(f"   Focal Length: {row['Focal Length']}")
        print(f"   T-Stop: {row['T-Stop']}")
        print()
    
    return needs_review

def analyze_low_confidence_patterns(df):
    """Analyze patterns in low-confidence lenses"""
    low_conf = df[df['Confidence Score'] < 0.5].copy()
    
    if len(low_conf) == 0:
        print("No lenses with very low confidence!")
        return
    
    print(f"\n=== LOW CONFIDENCE PATTERNS ({len(low_conf)} lenses) ===")
    
    # Check for missing manufacturers
    missing_manufacturer = low_conf[low_conf['Manufacturer'].isna() | (low_conf['Manufacturer'] == '')]
    if len(missing_manufacturer) > 0:
        print(f"\nMissing manufacturers ({len(missing_manufacturer)} lenses):")
        for _, row in missing_manufacturer.head(5).iterrows():
            print(f"  - {row['Original Name']}")
    
    # Check for missing series
    missing_series = low_conf[low_conf['Series'].isna() | (low_conf['Series'] == '')]
    if len(missing_series) > 0:
        print(f"\nMissing series ({len(missing_series)} lenses):")
        for _, row in missing_series.head(5).iterrows():
            print(f"  - {row['Original Name']}")
    
    # Check for missing T-stops
    missing_tstop = low_conf[low_conf['T-Stop'].isna() | (low_conf['T-Stop'] == '')]
    if len(missing_tstop) > 0:
        print(f"\nMissing T-stops ({len(missing_tstop)} lenses):")
        for _, row in missing_tstop.head(5).iterrows():
            print(f"  - {row['Original Name']}")

def suggest_improvements(df):
    """Suggest improvements based on analysis"""
    print(f"\n=== SUGGESTED IMPROVEMENTS ===")
    
    # Find common patterns in low-confidence lenses
    low_conf = df[df['Confidence Score'] < 0.6]
    
    if len(low_conf) == 0:
        print("No improvements needed!")
        return
    
    # Look for common words that might be manufacturers
    all_text = ' '.join(low_conf['Original Name'].astype(str)).lower()
    words = all_text.split()
    
    # Find potential manufacturers (words that appear multiple times)
    from collections import Counter
    word_counts = Counter(words)
    
    potential_manufacturers = []
    for word, count in word_counts.most_common(20):
        if count >= 3 and len(word) > 2 and word.isalpha():
            potential_manufacturers.append(word)
    
    if potential_manufacturers:
        print("Potential new manufacturers to add:")
        for manufacturer in potential_manufacturers[:10]:
            print(f"  - {manufacturer.title()}")
    
    # Look for common series patterns
    print(f"\nCommon patterns in low-confidence lenses:")
    for _, row in low_conf.head(10).iterrows():
        name = row['Original Name']
        if 'zoom' in name.lower():
            print(f"  - Zoom lens pattern: {name}")
        if '-' in name and 'mm' in name:
            print(f"  - Zoom focal length pattern: {name}")

def create_improved_parser():
    """Create an improved parser with custom patterns"""
    print(f"\n=== CREATING IMPROVED PARSER ===")
    
    # Load current parser
    parser = SimpleLensParser()
    
    # Add some common missing patterns
    custom_improvements = {
        'manufacturers': {
            'arri': ['arri'],
            'red': ['red'],
            'blackmagic': ['blackmagic'],
            'panasonic': ['panasonic'],
            'olympus': ['olympus'],
            'pentax': ['pentax'],
            'minolta': ['minolta'],
            'konica': ['konica'],
            'yashica': ['yashica'],
            'mamiya': ['mamiya'],
            'hasselblad': ['hasselblad'],
            'bronica': ['bronica'],
            'fuji': ['fuji'],
            'ricoh': ['ricoh'],
            'contax': ['contax']
        },
        'series': {
            'cinema zoom': ['cinema zoom'],
            'cine zoom': ['cine zoom'],
            'broadcast zoom': ['broadcast zoom'],
            'eng zoom': ['eng zoom'],
            'efp zoom': ['efp zoom'],
            'studio zoom': ['studio zoom'],
            'field zoom': ['field zoom'],
            'portrait': ['portrait'],
            'macro': ['macro'],
            'fisheye': ['fisheye'],
            'tilt-shift': ['tilt-shift', 'tilt shift'],
            'lensbaby': ['lensbaby'],
            'holga': ['holga'],
            'diana': ['diana'],
            'lomo': ['lomo']
        }
    }
    
    # Add custom patterns to parser
    for manufacturer, aliases in custom_improvements['manufacturers'].items():
        if manufacturer not in parser.manufacturers:
            parser.manufacturers[manufacturer] = aliases
    
    for series, patterns in custom_improvements['series'].items():
        if series not in parser.series_patterns:
            parser.series_patterns[series] = patterns
    
    print("Added custom patterns for:")
    print("  - 15 additional manufacturers")
    print("  - 15 additional series types")
    
    return parser

def save_improved_parser(parser, filename="improved_lens_parser.py"):
    """Save the improved parser to a file"""
    print(f"\nSaving improved parser to {filename}...")
    
    # This would require more complex code generation
    # For now, just save the custom patterns
    custom_patterns = {
        'manufacturers': parser.manufacturers,
        'series': parser.series_patterns
    }
    
    with open("custom_patterns.json", "w") as f:
        json.dump(custom_patterns, f, indent=2)
    
    print("Custom patterns saved to custom_patterns.json")

def main():
    """Main review function"""
    print("Lens Review and Improvement Tool")
    print("=" * 40)
    
    # Load parsed data
    df = load_parsed_data()
    if df is None:
        return
    
    # Show lenses needing review
    needs_review = show_lenses_needing_review(df, limit=15)
    
    # Analyze patterns
    analyze_low_confidence_patterns(df)
    
    # Suggest improvements
    suggest_improvements(df)
    
    # Create improved parser
    improved_parser = create_improved_parser()
    
    # Save improvements
    save_improved_parser(improved_parser)
    
    print(f"\n=== NEXT STEPS ===")
    print("1. Review the lenses shown above")
    print("2. Check custom_patterns.json for suggested additions")
    print("3. Manually correct any obvious errors in parsed_lenses_output.csv")
    print("4. Re-run process_existing_data.py with improved patterns")
    print("5. Consider adding more specific patterns for your lens types")

if __name__ == "__main__":
    main() 