#!/usr/bin/env python3
"""
Debug Innovision Parsing Issue
==============================

Debug script to investigate why "Innovision Probe II Plus T6.3 - LOW ANGLE PRISM"
is being incorrectly parsed with "Angenieux" as the manufacturer.
"""

from simple_lens_parser import SimpleLensParser

def debug_innovision_parsing():
    """Debug the Innovision parsing issue"""
    parser = SimpleLensParser()
    
    test_lens = "Innovision Probe II Plus T6.3 - LOW ANGLE PRISM"
    
    print("Debugging Innovision Parsing Issue")
    print("=" * 50)
    print(f"Test lens: '{test_lens}'")
    print()
    
    # Test preprocessing
    preprocessed = parser.preprocess_text(test_lens)
    print(f"Preprocessed text: '{preprocessed}'")
    print()
    
    # Test manufacturer detection
    print("--- Testing Manufacturer Detection ---")
    text = preprocessed
    best_match = None
    best_score = 0
    
    for manufacturer, aliases in parser.manufacturers.items():
        for alias in aliases:
            if alias in text:
                score = len(alias) / len(text) * 100
                print(f"  Found potential match for '{manufacturer}' with alias '{alias}' (Score: {score:.2f})")
                if score > best_score:
                    best_score = score
                    best_match = manufacturer

    final_manufacturer, final_score = parser.identify_manufacturer(text)
    print(f"\nFinal detected manufacturer: '{final_manufacturer}' (Score: {final_score:.2f})\n")

    # Test full parsing
    parsed = parser.parse_lens_name(test_lens)
    print("--- Full Parsing Result ---")
    print(f"  Manufacturer: '{parsed.manufacturer}'")
    print(f"  Series: '{parsed.series}'")
    print(f"  Focal Length: '{parsed.focal_length}'")
    print(f"  T-Stop: '{parsed.t_stop}'")
    print(f"  Lens Type: '{parsed.lens_type}'")
    print(f"  Confidence: {parsed.confidence_score:.3f}")

if __name__ == "__main__":
    debug_innovision_parsing() 