import pandas as pd

# 读取reference1.xlsx文件
df = pd.read_excel('reference1.xlsx')

# 打印列名
print("Reference1.xlsx 列名:")
print(df.columns.tolist())

# 打印前几行数据
print("\nReference1.xlsx 前5行数据:")
print(df.head())

# 读取config.py中定义的列名
print("\nConfig.py中定义的MATCHED_COLUMNS:")
from config import MATCHED_COLUMNS, MATERIAL_CODE_COLUMN
print(MATCHED_COLUMNS)
print(f"\nMATERIAL_CODE_COLUMN: {MATERIAL_CODE_COLUMN}")

# 检查reference1.xlsx中是否存在MATERIAL_CODE_COLUMN
print(f"\n{MATERIAL_CODE_COLUMN} 是否存在于reference1.xlsx: {MATERIAL_CODE_COLUMN in df.columns}")

# 如果不存在，打印所有可能的列名
if MATERIAL_CODE_COLUMN not in df.columns:
    print("\n可能的匹配列名:")
    for col in df.columns:
        print(f"  - {col}")