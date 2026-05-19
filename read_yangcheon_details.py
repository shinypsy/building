import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = '양천구_건축허가 및 사용승인현황_20260311.csv'
df = pd.read_csv(file_path, encoding='cp949')

print(df[['연번', '건축구분', '대지위치', '주용도']].head(20))
