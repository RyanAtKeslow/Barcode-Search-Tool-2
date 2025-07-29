#!/usr/bin/env python3
"""
Process Existing Lens Data
==========================

This script processes the corrected parsed_lenses_output.csv file using the simple lens parser
to improve the parsing while preserving manual corrections.
"""

import pandas as pd
from simple_lens_parser import SimpleLensParser
import logging
from pathlib import Path

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    """Main function to process existing data"""
    print("Processing Corrected Lens Data")
    print("=" * 50)
    
    # File paths
    corrected_lens_file = "parsed_lenses_output.csv"
    manual_edits_file = "Manual Edits.csv"
    output_file = "parsed_lenses_output_improved.csv"
    
    # Load corrected data
    print(f"Loading corrected data: {corrected_lens_file}")
    try:
        df = pd.read_csv(corrected_lens_file)
        print(f"Loaded {len(df)} corrected lens entries")
    except FileNotFoundError:
        print(f"Error: {corrected_lens_file} not found!")
        return
    
    # Load manual edits to check for deleted rows
    print(f"Loading manual edits: {manual_edits_file}")
    try:
        manual_df = pd.read_csv(manual_edits_file)
        print(f"Loaded {len(manual_df)} manual edit entries")
        
        # Create a set of original names from manual edits for comparison
        manual_original_names = set(manual_df['Original Name'].astype(str))
        print(f"Manual edits contains {len(manual_original_names)} unique original names")
    except FileNotFoundError:
        print(f"Warning: {manual_edits_file} not found! Proceeding without manual edits check.")
        manual_original_names = set()
    
    # Initialize parser
    parser = SimpleLensParser()
    improved_lenses = []
    
    print("Improving parsing for each lens...")
    
    for idx, row in df.iterrows():
        if (idx + 1) % 100 == 0:
            print(f"Processed {idx + 1}/{len(df)} lenses...")
        
        original_name = str(row['Original Name'])
        
        # Skip rows that were deleted from Manual Edits
        if manual_original_names and original_name not in manual_original_names:
            print(f"Skipping deleted row: {original_name}")
            continue
        
        parsed = parser.parse_lens_name(original_name)
        
        improved_row = {
            # Use parser output for all fields, allowing it to override manual corrections
            'Manufacturer': parsed.manufacturer,
            'Series': parsed.series,
            'Focal Length': parsed.focal_length,
            'T-Stop': parsed.t_stop,
            'Prime / Zoom / Special': parsed.lens_type,
            'Format': parsed.format,
            'Mount': parsed.mount,
            'Anamorphic / Spherical': parsed.anamorphic_spherical,
            'Anamorphic Squeeze Factor': parsed.anamorphic_squeeze,
            'Anamorphic Location': '',  # Never fill this field as per user request
            'Housing': parsed.housing,
            # Preserve all other fields as they were
            'Front Diameter (mm)': row.get('Front Diameter (mm)', ''),
            'Close Focus': row.get('Close Focus', ''),
            'Length (in)': row.get('Length (in)', ''),
            'Film Compatibility': row.get('Film Compatibility', ''),
            'Image Circle (mm)': row.get('Image Circle (mm)', ''),
            'Iris Blade Count': row.get('Iris Blade Count', ''),
            'Extender': row.get('Extender', ''),
            'LDS': row.get('LDS', ''),
            'i/Data': row.get('i/Data', ''),
            'Support Recommended': row.get('Support Recommended', ''),
            'Support Post Length (mm)': row.get('Support Post Length (mm)', ''),
            'Weight (lbs)': row.get('Weight (lbs)', ''),
            'Manufacture Year': row.get('Manufacture Year', ''),
            'Expander': row.get('Expander', ''),
            'Heden Motor Size': row.get('Heden Motor Size', ''),
            'Size': row.get('Size', ''),
            'Notes': parsed.notes,
            'Use Case': parsed.use_case,  # Column AC
            'Look': parsed.look,  # Column AD
            'Bokeh': row.get('Bokeh', ''),
            'Flare': parsed.flare,
            'Focus Falloff': row.get('Focus Falloff', ''),
            'Breathing': row.get('Breathing', ''),
            'Focus Scale': row.get('Focus Scale', ''),
            'Original Name': original_name,
            # Use parser confidence and needs review
            'Needs Review': parsed.needs_review,
            'Confidence Score': parsed.confidence_score
        }
        improved_lenses.append(improved_row)
    
    # Create DataFrame and save
    improved_df = pd.DataFrame(improved_lenses)
    improved_df.to_csv(output_file, index=False)
    
    print(f"\nImproved parsing complete!")
    print(f"Results saved to: {output_file}")
    
    # Generate summary
    generate_summary(improved_df, df)

def generate_summary(improved_df, original_df):
    """Generates and prints a summary of the improvement."""
    print(f"\n=== IMPROVEMENT SUMMARY ===")
    print(f"Total lenses processed: {len(improved_df)}")
    print(f"Average confidence: {improved_df['Confidence Score'].mean():.3f}")
    print(f"Lenses needing review: {len(improved_df[improved_df['Needs Review'] == True])}")
    
    # Compare with original
    original_needs_review = len(original_df[original_df['Needs Review'] == True])
    improved_needs_review = len(improved_df[improved_df['Needs Review'] == True])
    improvement = original_needs_review - improved_needs_review
    
    print(f"\n=== COMPARISON WITH ORIGINAL ===")
    print(f"Original lenses needing review: {original_needs_review}")
    print(f"Improved lenses needing review: {improved_needs_review}")
    print(f"Improvement: {improvement} fewer lenses need review")
    
    # Show manufacturer distribution
    print(f"\n=== MANUFACTURER DISTRIBUTION ===")
    manufacturer_counts = improved_df['Manufacturer'].value_counts()
    for manufacturer, count in manufacturer_counts.head(10).items():
        if pd.notna(manufacturer) and manufacturer != "":
            print(f"  {manufacturer}: {count}")
    
    # Show series distribution
    print(f"\n=== SERIES DISTRIBUTION ===")
    series_counts = improved_df['Series'].value_counts()
    for series, count in series_counts.head(10).items():
        if pd.notna(series) and series != "":
            print(f"  {series}: {count}")
    
    # Show format distribution
    print(f"\n=== FORMAT DISTRIBUTION ===")
    format_counts = improved_df['Format'].value_counts()
    for format_type, count in format_counts.head(10).items():
        if pd.notna(format_type) and format_type != "":
            print(f"  {format_type}: {count}")
    
    # Show mount distribution
    print(f"\n=== MOUNT DISTRIBUTION ===")
    mount_counts = improved_df['Mount'].value_counts()
    for mount, count in mount_counts.head(10).items():
        if pd.notna(mount) and mount != "":
            print(f"  {mount}: {count}")
    
    # Show high-confidence examples
    print(f"\n=== HIGH-CONFIDENCE EXAMPLES ===")
    high_conf = improved_df[improved_df['Confidence Score'] >= 0.8]
    for _, row in high_conf.head(5).iterrows():
        print(f"  {row['Original Name']} -> {row['Manufacturer']} {row['Series']} {row['Focal Length']} {row['T-Stop']}")
    
    # Show lenses that still need review
    print(f"\n=== LENSES STILL NEEDING REVIEW ===")
    needs_review = improved_df[improved_df['Needs Review'] == True]
    for _, row in needs_review.head(10).iterrows():
        print(f"  {row['Original Name']} (Confidence: {row['Confidence Score']:.3f})")
    
    # Save a summary report
    summary_file = "improvement_summary.txt"
    with open(summary_file, 'w') as f:
        f.write("Lens Parsing Improvement Summary Report\n")
        f.write("=" * 40 + "\n\n")
        f.write(f"Total lenses processed: {len(improved_df)}\n")
        f.write(f"Average confidence: {improved_df['Confidence Score'].mean():.3f}\n")
        f.write(f"Lenses needing review: {len(improved_df[improved_df['Needs Review'] == True])}\n")
        f.write(f"Improvement: {improvement} fewer lenses need review\n\n")
        
        f.write("Top Manufacturers:\n")
        for manufacturer, count in manufacturer_counts.head(10).items():
            if pd.notna(manufacturer) and manufacturer != "":
                f.write(f"  {manufacturer}: {count}\n")
        
        f.write("\nTop Series:\n")
        for series, count in series_counts.head(10).items():
            if pd.notna(series) and series != "":
                f.write(f"  {series}: {count}\n")
        
        f.write("\nFormat Distribution:\n")
        for format_type, count in format_counts.head(10).items():
            if pd.notna(format_type) and format_type != "":
                f.write(f"  {format_type}: {count}\n")
        
        f.write("\nMount Distribution:\n")
        for mount, count in mount_counts.head(10).items():
            if pd.notna(mount) and mount != "":
                f.write(f"  {mount}: {count}\n")
    
    print(f"\nImprovement summary saved to: {summary_file}")
    print(f"Improved results saved to: parsed_lenses_output_improved.csv")
    print(f"\nYou can now replace the original parsed_lenses_output.csv with parsed_lenses_output_improved.csv")

if __name__ == "__main__":
    main() 