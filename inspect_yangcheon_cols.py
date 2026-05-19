import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = '양천구_건축허가 및 사용승인현황_20260311.xlsx'
try:
    df = pd.read_excel(file_path)

    print("CSV 로딩 성공!")
    print(f"Columns: {df.columns.tolist()}")
    print("\n상위 3행 연면적 샘플:")
    # 연면적이 포함된 컬럼 찾기
    area_cols = [col for col in df.columns if '연면적' in col]
    print(f"연면적 관련 컬럼들: {area_cols}")
    if area_cols:
        print(df[['대지위치', '건물명'] + area_cols].head(5))
except Exception as e:
    print(f"Error: {e}")
