import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = '시설관리 업체 순위_정보포함.xlsx'
df = pd.read_excel(file_path)

print("Columns in final file:")
print(df.columns.tolist())

print("\nShape:", df.shape)

print("\nFirst 10 rows:")
print(df.head(10))

# Count N/As
na_count = (df == 'N/A').sum()
print("\nN/A values per column:")
print(na_count)

success_rate = (len(df) - na_count['대표자']) / len(df) * 100
print(f"\nScraping Success Rate: {success_rate:.2f}%")
