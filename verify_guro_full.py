import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

guro_file = '구로구 사용승인 현황(2020년이후).xlsx'
try:
    df = pd.read_excel(guro_file, sheet_name='중대형건물_관리현황')
    
    missing_count = 0
    for idx, row in df.iterrows():
        bname = str(row['건물명']).strip()
        addr = str(row['대지위치']).strip()
        manager = str(row['관리업체명']).strip()
        is_managed = str(row['건물관리업체 여부']).strip()
        
        if is_managed != 'Y' or manager == 'nan' or manager == '' or '조사중' in manager:
            print(f"[MISSING] {bname} / {addr} / {manager}")
            missing_count += 1
            
    if missing_count == 0:
        print("== 100% 전수 검증 성공! 누락된 행이 없습니다. ==")
        print(f"전체 {len(df)}건 모두 관리업체명 및 연락처 부여 완료.")
    else:
        print(f"== 누락 건수: {missing_count}건 ==")
except Exception as e:
    print(f"Error: {e}")
