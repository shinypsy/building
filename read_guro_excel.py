import pandas as pd
import sys
file_path = r'd:/Dev/Project/sample/구로구 사용승인 현황(2020년 이후).xlsx'
try:
    xl = pd.ExcelFile(file_path)
    print('시트 리스트:', xl.sheet_names)
    if '중대형건물_관리현황' in xl.sheet_names:
        df = xl.parse('중대형건물_관리현황')
        print('데이터 미리보기 (head):')
        print(df.head().to_string(index=False))
    else:
        print('시트 없음')
except Exception as e:
    print('오류:', e)
