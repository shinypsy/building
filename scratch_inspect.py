import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

yangcheon_file = '양천구_건축허가 및 사용승인현황_20260311.xlsx'
guro_file = '구로구 사용승인 현황(2020년이후).xlsx'

print("== Yangcheon-gu Sheets ==")
try:
    xl_y = pd.ExcelFile(yangcheon_file)
    print(xl_y.sheet_names)
    if '중대형건물_관리현황' in xl_y.sheet_names:
        df_y = pd.read_excel(yangcheon_file, sheet_name='중대형건물_관리현황', nrows=5)
        print("Columns in '중대형건물_관리현황':")
        print(df_y.columns.tolist())
except Exception as e:
    print("Error reading Yangcheon-gu:", e)

print("\n== Guro-gu Sheets ==")
try:
    xl_g = pd.ExcelFile(guro_file)
    print(xl_g.sheet_names)
    for sheet in xl_g.sheet_names:
        df_g = pd.read_excel(guro_file, sheet_name=sheet, nrows=5)
        print(f"Columns in '{sheet}':")
        print(df_g.columns.tolist())
except Exception as e:
    print("Error reading Guro-gu:", e)
