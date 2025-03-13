import pandas as pd
import os
import argparse

# Default configuration
PRESERVED_COLUMNS = ['Column1', 'Column2', 'Column3']
MATERIAL_CODE_COLUMN = 'MaterialCode'
MATCHED_COLUMNS = ['MatchedColumn1', 'MatchedColumn2']
FIXED_COLUMNS = {'FixedColumn1': 'Fixed Value 1', 'FixedColumn2': 'Fixed Value 2'}

# Try importing the configuration
try:
    from config import PRESERVED_COLUMNS, MATERIAL_CODE_COLUMN, MATCHED_COLUMNS, FIXED_COLUMNS
except ImportError:
    print("Warning: config.py file not found. Using default configuration.")

def convert_excel(input_file, reference_file, output_file):
    """
    Convert Excel file according to specified requirements.
    
    Args:
        input_file (str): Path to the first Excel file (source data)
        reference_file (str): Path to the reference Excel file (for material code matching)
        output_file (str): Path to save the output Excel file
    """
    # Read the input Excel file
    print(f"Reading input file: {input_file}")
    df_input = pd.read_excel(input_file)
    
    # Read the reference Excel file
    print(f"Reading reference file: {reference_file}")
    df_reference = pd.read_excel(reference_file)
    
    # Create a new DataFrame for the output
    df_output = pd.DataFrame()
    
    # Copy preserved columns (green headers) from input file
    for col in PRESERVED_COLUMNS:
        if col in df_input.columns:
            df_output[col] = df_input[col]
        else:
            print(f"Warning: Preserved column '{col}' not found in input file")
    
    # Match columns by material code (yellow headers)
    if MATERIAL_CODE_COLUMN in df_input.columns and MATERIAL_CODE_COLUMN in df_reference.columns:
        # Create a mapping dictionary for faster lookups
        reference_dict = {}
        for col in MATCHED_COLUMNS:
            if col in df_reference.columns:
                reference_dict[col] = df_reference.set_index(MATERIAL_CODE_COLUMN)[col].to_dict()
            else:
                print(f"Warning: Matched column '{col}' not found in reference file")
        
        # Add matched columns to output DataFrame
        for col in MATCHED_COLUMNS:
            if col in reference_dict:
                df_output[col] = df_input[MATERIAL_CODE_COLUMN].map(reference_dict[col])
    else:
        print(f"Warning: Material code column '{MATERIAL_CODE_COLUMN}' not found in one of the files")
    
    # Add fixed columns
    for col, value in FIXED_COLUMNS.items():
        df_output[col] = value
    
    # Save the output Excel file
    print(f"Saving output file: {output_file}")
    df_output.to_excel(output_file, index=False)
    print("Conversion completed successfully!")
    return df_output

def main():
    parser = argparse.ArgumentParser(description='Convert Excel files according to specified format')
    parser.add_argument('--input', required=True, help='Path to the input Excel file')
    parser.add_argument('--reference', required=True, help='Path to the reference Excel file')
    parser.add_argument('--output', required=True, help='Path to save the output Excel file')
    
    args = parser.parse_args()
    
    convert_excel(args.input, args.reference, args.output)

if __name__ == "__main__":
    main()