import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

guro_file = '구로구 사용승인 현황(2020년이후).xlsx'

try:
    df = pd.read_excel(guro_file, sheet_name='중대형건물_관리현황')
    print("== 중대형건물_관리현황 시트 검증 ==")
    print(f"행 수: {len(df)}")
    print(f"컬럼 목록 (총 {len(df.columns)}개):")
    print(df.columns.tolist())
    print("샘플 데이터 (첫 2건):")
    print(df.head(2).to_dict(orient='records'))
except Exception as e:
    print(f"오류: {e}")
