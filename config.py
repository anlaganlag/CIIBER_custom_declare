# Configuration for Excel converter

# Columns to preserve from the input file (green headers)
PRESERVED_COLUMNS = [
    'DESCRIPTION',
    'Qty',
    'Amount'
    # Add more columns as needed
]

# Column containing the material code for matching
MATERIAL_CODE_COLUMN = 'Material code'

# Columns to match from the reference file (yellow headers)
MATCHED_COLUMNS = [
    '商品编号',
    # Add more columns as needed
]

# Fixed columns to add to the output file
FIXED_COLUMNS = {
    '币值1': '美元',
    '币值2': '人民币',
    '币值3': 'BTC',

    # Add more columns as needed
}

# Output file name
OUTPUT_FILE_NAME = 'output.xlsx'    