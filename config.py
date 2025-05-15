# Configuration for Excel converter

# Columns to preserve from the input file (green headers)
PRESERVED_COLUMNS = [
    '项号',
    '商品名称',
    '型号',
    '数量',
    '单位',
    '单价',
    '总价',
    '净重',
    # Add more columns as needed
]

# Column containing the material code for matching
MATERIAL_CODE_COLUMN = 'Part Number'

# Columns to match from the reference file (yellow headers)
MATCHED_COLUMNS = [
    '商品编码',
    '申报要素',
    # Add more columns as needed
]

# Fixed columns to add to the output file
FIXED_COLUMNS = {
    '币制': '美元',
    '原产国（地区）': '中国',
    '最终目的国（地区）': '印度',
    '境内货源地': '深圳特区',
    '征免': '照章征税',
    # Add more columns as needed
}

# Output file name
OUTPUT_FILE_NAME = 'output.xlsx'