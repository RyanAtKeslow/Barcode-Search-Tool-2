import pandas as pd
import re
import numpy as np

def normalize_text(value):
    """Convert all text to lowercase and handle NaN values"""
    if pd.isna(value) or value == '':
        return ''
    return str(value).lower().strip()

def normalize_focal_length(value):
    """Normalize focal length to remove 'mm' suffix and lowercase"""
    if pd.isna(value) or value == '':
        return ''
    
    # Convert to string and lowercase
    focal = str(value).lower().strip()
    
    # Remove 'mm' suffix if present
    focal = re.sub(r'mm$', '', focal)
    
    return focal.strip()

def normalize_t_stop(value):
    """Normalize T-stop to remove 'T' prefix and lowercase"""
    if pd.isna(value) or value == '':
        return ''
    
    # Convert to string and lowercase
    t_stop = str(value).lower().strip()
    
    # Remove 'T' or 't' prefix if present
    t_stop = re.sub(r'^t', '', t_stop)
    
    return t_stop.strip()

def normalize_measurement_with_unit(value, unit_suffix):
    """Normalize measurements by removing unit suffix (like 'mm', 'lbs')"""
    if pd.isna(value) or value == '':
        return ''
    
    # Convert to string and lowercase
    measurement = str(value).lower().strip()
    
    # Remove the unit suffix if present
    pattern = rf'{unit_suffix}$'
    measurement = re.sub(pattern, '', measurement)
    
    return measurement.strip()

def normalize_diameter(value):
    """Normalize front diameter measurements"""
    return normalize_measurement_with_unit(value, 'mm')

def normalize_weight(value):
    """Normalize weight measurements"""
    # Handle both 'lbs' and 'lb'
    if pd.isna(value) or value == '':
        return ''
    
    weight = str(value).lower().strip()
    weight = re.sub(r'\s*lbs?$', '', weight)  # Remove 'lb' or 'lbs'
    
    return weight.strip()

def normalize_close_focus(value):
    """Normalize close focus measurements (remove quotes)"""
    if pd.isna(value) or value == '':
        return ''
    
    # Convert to string and lowercase
    focus = str(value).lower().strip()
    
    # Remove quotes from measurements like 8", 2' 3", etc.
    focus = re.sub(r'["\']', '', focus)
    
    return focus.strip()

def normalize_length(value):
    """Normalize length measurements (remove quotes)"""
    return normalize_close_focus(value)  # Same logic as close focus

def normalize_image_circle(value):
    """Normalize image circle measurements"""
    if pd.isna(value) or value == '':
        return ''
    
    # Convert to string and lowercase
    circle = str(value).lower().strip()
    
    # Handle complex descriptions like "31mm clear / 38mm with falloff"
    # Remove 'mm' suffix but keep the descriptive text
    circle = re.sub(r'mm(?=\s|$)', '', circle)
    
    return circle.strip()

def normalize_iris_blade_count(value):
    """Normalize iris blade count (remove parenthetical descriptions)"""
    if pd.isna(value) or value == '':
        return ''
    
    # Convert to string and lowercase
    iris = str(value).lower().strip()
    
    # Extract just the number, remove descriptions like "(triangle)"
    iris = re.sub(r'\s*\([^)]*\)', '', iris)
    
    return iris.strip()

def normalize_csv_file(file_path, output_path):
    """Normalize a single CSV file"""
    print(f"Processing {file_path}...")
    
    # Read the CSV file
    df = pd.read_csv(file_path)
    
    # Get column names
    cols = df.columns.tolist()
    
    # Apply normalization to each column based on its content
    for col in cols:
        if col.lower() == 'focal length':
            df[col] = df[col].apply(normalize_focal_length)
        elif col.lower() == 't-stop':
            df[col] = df[col].apply(normalize_t_stop)
        elif 'diameter' in col.lower():
            df[col] = df[col].apply(normalize_diameter)
        elif 'close focus' in col.lower():
            df[col] = df[col].apply(normalize_close_focus)
        elif 'length' in col.lower() and '(' in col.lower():  # "Length (in)"
            df[col] = df[col].apply(normalize_length)
        elif 'weight' in col.lower():
            df[col] = df[col].apply(normalize_weight)
        elif 'image circle' in col.lower():
            df[col] = df[col].apply(normalize_image_circle)
        elif 'iris blade count' in col.lower():
            df[col] = df[col].apply(normalize_iris_blade_count)
        else:
            # For all other text columns, just lowercase and clean
            df[col] = df[col].apply(normalize_text)
    
    # Save the normalized file
    df.to_csv(output_path, index=False)
    print(f"Normalized file saved to {output_path}")
    
    return df

def main():
    """Main function to normalize both CSV files"""
    
    # File paths
    core_input = "Lens Database/Final Flatten/Core.csv"
    tech_input = "Lens Database/Final Flatten/Tech Inf.csv"
    
    core_output = "Lens Database/Final Flatten/Core_normalized.csv"
    tech_output = "Lens Database/Final Flatten/Tech_Inf_normalized.csv"
    
    print("Starting CSV normalization process...")
    print("=" * 50)
    
    # Normalize both files
    try:
        core_df = normalize_csv_file(core_input, core_output)
        tech_df = normalize_csv_file(tech_input, tech_output)
        
        print("\n" + "=" * 50)
        print("Normalization completed successfully!")
        print(f"Core.csv: {len(core_df)} rows processed")
        print(f"Tech Inf.csv: {len(tech_df)} rows processed")
        
        # Show sample of normalized data
        print("\nSample of normalized Core data:")
        print(core_df[['Manufacturer', 'Focal Length', 'T-Stop']].head(3))
        
        print("\nSample of normalized Tech Inf data:")
        print(tech_df[['Manufacturer', 'Focal Length', 'T-Stop']].head(3))
        
    except FileNotFoundError as e:
        print(f"Error: File not found - {e}")
    except Exception as e:
        print(f"Error during processing: {e}")

if __name__ == "__main__":
    main()