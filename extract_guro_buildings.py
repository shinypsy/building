import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

guro_file = '구로구 사용승인 현황(2020년이후).xlsx'
try:
    df = pd.read_excel(guro_file, sheet_name='중대형건물_관리현황')
    print("== Guro-gu Buildings ==")
    for idx, row in df.iterrows():
        print(f"- {row['건물명']} / {row['대지위치']}")
except Exception as e:
    print(f"Error: {e}")
