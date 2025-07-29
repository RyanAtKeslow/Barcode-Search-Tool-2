# Lens Database - Machine Learning Module

This directory contains a Machine Learning module for parsing lens name strings into structured data.

## Quick Start

### 1. Navigate to the Machine Learning directory
```bash
cd "Machine Learning"
```

### 2. Set up the environment
```bash
python setup.py
```

### 3. Test the parser
```bash
python test_parser.py
```

### 4. Process your lens data
```bash
python lens_parser_ml.py input.csv output.csv
```

### 5. Run the example with existing data
```bash
python example_usage.py
```

## Files in the Machine Learning Module

- **`lens_parser_ml.py`** - Main parsing script with NLP and ML capabilities
- **`train_lens_parser.py`** - Training and validation script
- **`test_parser.py`** - Test script to demonstrate functionality
- **`example_usage.py`** - Example showing integration with existing lens database
- **`setup.py`** - Automated setup script for dependencies
- **`requirements.txt`** - Python package dependencies
- **`README.md`** - Comprehensive documentation
- **`run_example.py`** - Launcher script for demonstrations

## Integration with Existing Data

The Machine Learning module is designed to work with your existing lens database files:

- **`ESC Raw Lenses.csv`** - Raw lens name strings (input)
- **`ESC_Raw_Lenses_Flat.csv`** - Existing parsed data (for comparison)
- **`Flattened_Lens_Inventory.csv`** - Complete lens inventory

## Example Usage

### Process existing raw lens data:
```bash
cd "Machine Learning"
python example_usage.py
```

This will:
1. Process the existing `ESC Raw Lenses.csv` file
2. Compare results with existing parsed data
3. Create an enhanced database with confidence scores
4. Generate detailed analysis reports

### Create training data:
```bash
cd "Machine Learning"
python train_lens_parser.py --create-sample training_data.csv
```

### Validate parser accuracy:
```bash
cd "Machine Learning"
python train_lens_parser.py training_data.csv --validate-only
```

## Output Files

The parser generates several output files:

- **`parsed_lenses_output.csv`** - New parsed results
- **`enhanced_lens_database.csv`** - Enhanced existing database with confidence scores
- **`training_report.json`** - Detailed training analysis
- **`validation_results.csv`** - Validation results with predictions vs ground truth

## Features

- **Advanced NLP parsing** with confidence scoring
- **Fuzzy matching** for manufacturer and series identification
- **Quality control** with "needs review" flags
- **Training support** for improving accuracy
- **Integration** with existing lens database structure

## Requirements

- Python 3.7 or higher
- See `Machine Learning/requirements.txt` for package dependencies
- spaCy English language model

## Support

For detailed documentation and troubleshooting, see `Machine Learning/README.md`. 