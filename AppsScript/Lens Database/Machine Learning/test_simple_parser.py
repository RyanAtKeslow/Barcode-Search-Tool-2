#!/usr/bin/env python3
"""
Simple Lens Parser Test
=======================

This script tests the simple lens parser with sample lens names.
No heavy dependencies required - just basic Python libraries.
"""

import pandas as pd
from simple_lens_parser import SimpleLensParser
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_individual_lenses():
    """Test the parser with individual lens names"""
    parser = SimpleLensParser(confidence_threshold=0.7)
    
    # Sample lens names to test
    test_lenses = [
        "12mm Master Prime T1.3",
        "Canon 6.6-66mm T2.5 Zoom",
        "Cooke S4/i 18mm T2.0",
        "Zeiss Ultra Prime 50mm T1.9 (LDS)",
        "Nikon 24-70mm f/2.8G ED",
        "Leica Summilux-C 35mm T1.4",
        "Fujinon MK 18-55mm T2.9",
        "Angenieux Optimo 24-290mm T2.8",
        "Hawk V-Lite 75mm T1.5 Anamorphic",
        "Sigma Cine FF High Speed 40mm T1.5"
    ]
    
    print("=== SIMPLE LENS PARSER TEST RESULTS ===\n")
    
    for lens_name in test_lenses:
        print(f"Input: {lens_name}")
        
        # Parse the lens
        parsed = parser.parse_lens_name(lens_name)
        
        # Display results
        print(f"  Manufacturer: {parsed.manufacturer}")
        print(f"  Series: {parsed.series}")
        print(f"  Focal Length: {parsed.focal_length}")
        print(f"  T-Stop: {parsed.t_stop}")
        print(f"  Type: {parsed.prime_zoom_special}")
        print(f"  Mount: {parsed.mount}")
        print(f"  Format: {parsed.format}")
        print(f"  Anamorphic: {parsed.anamorphic_spherical}")
        print(f"  LDS: {parsed.lds}")
        print(f"  Confidence: {parsed.confidence_score:.3f}")
        print(f"  Needs Review: {parsed.needs_review}")
        print()

def test_csv_processing():
    """Test processing a CSV file"""
    # Create a sample CSV file
    sample_data = [
        {"Lens Name": "12mm Master Prime T1.3"},
        {"Lens Name": "Canon 6.6-66mm T2.5 Zoom"},
        {"Lens Name": "Cooke S4/i 18mm T2.0"},
        {"Lens Name": "Zeiss Ultra Prime 50mm T1.9 (LDS)"},
        {"Lens Name": "Nikon 24-70mm f/2.8G ED"}
    ]
    
    # Save sample CSV
    df = pd.DataFrame(sample_data)
    input_file = "test_input.csv"
    output_file = "test_output.csv"
    df.to_csv(input_file, index=False)
    
    print(f"Created test input file: {input_file}")
    
    # Process the CSV
    parser = SimpleLensParser(confidence_threshold=0.7)
    parser.parse_csv(input_file, output_file)
    
    # Display results
    print(f"\nResults saved to: {output_file}")
    
    # Show summary
    results_df = pd.read_csv(output_file)
    print(f"\nProcessed {len(results_df)} lenses")
    print(f"Average confidence: {results_df['Confidence Score'].mean():.3f}")
    print(f"Lenses needing review: {len(results_df[results_df['Needs Review'] == True])}")
    
    # Show sample results
    print("\nSample results:")
    for _, row in results_df.head(3).iterrows():
        print(f"  {row['Original Name']} -> {row['Manufacturer']} {row['Series']} {row['Focal Length']} {row['T-Stop']}")

def test_with_existing_data():
    """Test with existing lens data"""
    print("\n=== TESTING WITH EXISTING DATA ===\n")
    
    # Check if existing data file exists
    import os
    existing_file = "../ESC Raw Lenses.csv"
    
    if os.path.exists(existing_file):
        print(f"Found existing data file: {existing_file}")
        
        # Read a sample of the data
        df = pd.read_csv(existing_file)
        print(f"Total rows in file: {len(df)}")
        
        # Take first 10 rows for testing
        sample_df = df.head(10)
        sample_file = "sample_existing_data.csv"
        sample_df.to_csv(sample_file, index=False)
        
        print(f"Created sample file: {sample_file}")
        
        # Process the sample
        parser = SimpleLensParser(confidence_threshold=0.7)
        output_file = "sample_parsed_results.csv"
        parser.parse_csv(sample_file, output_file)
        
        # Show results
        results_df = pd.read_csv(output_file)
        print(f"\nParsed {len(results_df)} lenses from existing data")
        print(f"Average confidence: {results_df['Confidence Score'].mean():.3f}")
        
        print("\nSample parsed results:")
        for _, row in results_df.head(5).iterrows():
            print(f"  {row['Original Name']} -> {row['Manufacturer']} {row['Series']} {row['Focal Length']} {row['T-Stop']}")
    else:
        print(f"Existing data file not found: {existing_file}")
        print("Skipping existing data test")

def main():
    """Run all tests"""
    print("Starting Simple Lens Parser Tests...\n")
    
    try:
        # Test individual lens parsing
        test_individual_lenses()
        
        # Test CSV processing
        test_csv_processing()
        
        # Test with existing data
        test_with_existing_data()
        
        print("\n=== ALL TESTS COMPLETED ===")
        print("Check the generated files for detailed results.")
        
    except Exception as e:
        logger.error(f"Test failed: {e}")
        print(f"\nError during testing: {e}")
        print("Please check that pandas is installed:")
        print("  pip3 install pandas")

if __name__ == "__main__":
    main() 