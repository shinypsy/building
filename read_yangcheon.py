import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = '양천구_건축허가 및 사용승인현황_20260311.csv'
print(f"Reading '{file_path}' to inspect structure...")

encodings = ['utf-8', 'cp949', 'euc-kr', 'utf-16']
df = None

for enc in encodings:
    try:
        df = pd.read_csv(file_path, encoding=enc)
        print(f"Successfully read with encoding: {enc}")
        break
    except Exception as e:
        print(f"Failed with {enc}: {e}")

if df is not None:
    print(f"Shape: {df.shape}")
    print("\nColumns:")
    print(df.columns.tolist())
    print("\nFirst 5 rows:")
    print(df.head(5))
else:
    print("Failed to read the file with all attempted encodings.")
