import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = '양천구_건축허가 및 사용승인현황_20260311.xlsx'
df = pd.read_excel(file_path)

# 연면적 조건 필터링
# 10000 <= 연면적 < 30000
target_df = df[(df['연면적(제곱미터)'] >= 10000) & (df['연면적(제곱미터)'] < 30000)]

print(f"연면적 1만~3만㎡ 필터링된 총 행 수: {len(target_df)}")

# 건물명 기준 중복 제거
unique_buildings = target_df.drop_duplicates(subset=['건물명'])
print(f"건물명 중복 제거 후 총 건물 수: {len(unique_buildings)}")

print("\n중복 제거 후 건물 리스트:")
for idx, row in unique_buildings.iterrows():
    print(f"건물명: {row['건물명']} | 대지위치: {row['대지위치']} | 연면적: {row['연면적(제곱미터)']}㎡ | 주용도: {row['주용도']}")
