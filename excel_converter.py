import pandas as pd
import os
import argparse

# Default configuration
PRESERVED_COLUMNS = ['Column1', 'Column2', 'Column3']
MATERIAL_CODE_COLUMN = 'MaterialCode'
MATCHED_COLUMNS = ['MatchedColumn1', 'MatchedColumn2']
FIXED_COLUMNS = {'FixedColumn1': 'Fixed Value 1', 'FixedColumn2': 'Fixed Value 2'}

# Column name mapping from English to Chinese
COLUMN_MAPPING = {
    'NO.': '项号',
    'DESCRIPTION': '品名',
    'Model NO.': '型号',
    'Qty': '数量',
    'Unit': '单位',
    'Amount': '总价',
    'net weight': '净重',
    'Unit Price': '单价',
}

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
    
    # Strip whitespace from column names
    df_input.columns = df_input.columns.str.strip()
    
    # Strip whitespace from string data in all columns
    for column in df_input.select_dtypes(include=['object']).columns:
        df_input[column] = df_input[column].str.strip()
    
    # 打印读取到的列名
    print(f"Input file columns: {df_input.columns.tolist()}")
    
    # Combine multi-level column headers
    
    # Read the reference Excel file
    print(f"Reading reference file: {reference_file}")
    df_reference = pd.read_excel(reference_file)   
    # Create a new DataFrame for the output
    df_output = pd.DataFrame()
    # Define the desired column order
    column_order = [
        '项号',
        '商品编号',
        '品名',
        '型号',
        '申报要素',
        '数量',
        '单位',
        '单价',
        '总价',
        '币制',
        '原产国（地区）',
        '最终目的国（地区）',
        '境内货源地',
        '征免',
        '净重'
    ]    
    # Copy preserved columns (green headers) from input file with Chinese column names
    for col in PRESERVED_COLUMNS:
        for eng_col, cn_col in COLUMN_MAPPING.items():
            if eng_col in df_input.columns and cn_col == col:
                df_output[col] = df_input[eng_col]
                break
        else:
            print(f"Warning: Column '{col}' not found in input file")
    
    # Match columns by material code (yellow headers)
    material_code_eng = next((k for k, v in COLUMN_MAPPING.items() if v == MATERIAL_CODE_COLUMN), MATERIAL_CODE_COLUMN)
    print(f"Looking for material code column: {material_code_eng} or {MATERIAL_CODE_COLUMN}")
    
    if material_code_eng in df_input.columns and MATERIAL_CODE_COLUMN in df_reference.columns:
        print("Found material code columns in both files")
        # Create a mapping dictionary for faster lookups
        reference_dict = {}
        for col in MATCHED_COLUMNS:
            if col in df_reference.columns:
                print(f"Creating mapping for column '{col}'")
                reference_dict[col] = df_reference.set_index(MATERIAL_CODE_COLUMN)[col].to_dict()
            else:
                print(f"Warning: Matched column '{col}' not found in reference file")
        
        # Add matched columns to output DataFrame
        for col in MATCHED_COLUMNS:
            if col in reference_dict:
                print(f"Applying mapping for column '{col}'")
                df_output[col] = df_input[material_code_eng].map(reference_dict[col])
    else:
        print(f"Warning: Material code column not found in one of the files")
        print(f"Input columns available: {df_input.columns.tolist()}")
        print(f"Reference columns available: {df_reference.columns.tolist()}")
    
    # Add fixed columns
    for col, value in FIXED_COLUMNS.items():
        print(f"Adding fixed column '{col}' with value '{value}'")
        df_output[col] = value
    
    # Reorder columns according to the desired order
    print("Reordering columns according to specified order")
    # Define column order using configuration
    column_order = PRESERVED_COLUMNS + MATCHED_COLUMNS + list(FIXED_COLUMNS.keys())
    print(f"Column order: {column_order}")
    df_output = df_output.reindex(columns=column_order)
    print(f"Final columns: {df_output.columns.tolist()}")
    
    # Save the output Excel file
    print(f"Saving output file: {output_file}")
    df_output.to_excel(output_file, index=False)
    print("Conversion completed successfully!")
    
    # Return the DataFrame for potential further processing
    return df_output

def main():
    parser = argparse.ArgumentParser(description='Convert Excel files according to specified format')
    parser.add_argument('input', help='Path to the input Excel file')
    parser.add_argument('reference', help='Path to the reference Excel file')
    parser.add_argument('output', help='Path to save the output Excel file')
    
    args = parser.parse_args()
    
    convert_excel(args.input, args.reference, args.output)

if __name__ == "__main__":
    main()