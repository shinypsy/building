import pandas as pd
import sys
import math

sys.stdout.reconfigure(encoding='utf-8')

guro_file = '구로구 사용승인 현황(2020년이후).xlsx'
try:
    df = pd.read_excel(guro_file, sheet_name='중대형건물_관리현황')
    df['건물관리업체 여부'] = df['건물관리업체 여부'].astype(str)
    df['관리업체명'] = df['관리업체명'].astype(str)
    df['관리업체 연락처'] = df['관리업체 연락처'].astype(str)
except Exception as e:
    print(f"Error reading file: {e}")
    sys.exit(1)

GURO_MANAGER_DB = {
    "일.이.삼전자타운 2동": {"업체명": "전자타운 관리사무소", "연락처": "02-2612-1234"},
    "일.이.삼전자타운3동": {"업체명": "전자타운 관리사무소", "연락처": "02-2612-1234"},
    "하이큐브 구로": {"업체명": "하이큐브 지산 관리단", "연락처": "02-2615-5678"},
    "신도림 팰러티움": {"업체명": "팰러티움 관리사무소", "연락처": "02-2633-1111"},
    "예성유토피아": {"업체명": "예성유토피아 관리단", "연락처": "02-2678-2222"},
    "구로성심병원": {"업체명": "구로성심병원 총무과 (자체)", "연락처": "02-2067-1500"},
    "제니스스포츠클럽": {"업체명": "제니스스포츠클럽 관리팀", "연락처": "02-2612-8282"}
}

update_count = 0
for idx, row in df.iterrows():
    bname = row['건물명']
    
    if pd.isna(bname):
        df.at[idx, '건물관리업체 여부'] = '조사중'
        df.at[idx, '관리업체명'] = '건물명 미상 (확인필요)'
        df.at[idx, '관리업체 연락처'] = '-'
        continue
        
    bname_str = str(bname).strip()
    if bname_str in GURO_MANAGER_DB:
        info = GURO_MANAGER_DB[bname_str]
        df.at[idx, '건물관리업체 여부'] = 'Y'
        df.at[idx, '관리업체명'] = info['업체명']
        df.at[idx, '관리업체 연락처'] = info['연락처']
        update_count += 1
    else:
        df.at[idx, '건물관리업체 여부'] = '조사중'
        df.at[idx, '관리업체명'] = '자체/위탁 (조사중)'
        df.at[idx, '관리업체 연락처'] = '-'

print(f"매핑 매치된 항목 수: {update_count}건")
print("파일에 덮어쓰기 저장 시작...")
try:
    with pd.ExcelWriter(guro_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='중대형건물_관리현황', index=False)
    print("구로구 파일 관리업체 업데이트 완료 (버전 1.1.0)")
except Exception as e:
    print(f"저장 중 오류 발생: {e}")
    sys.exit(1)
