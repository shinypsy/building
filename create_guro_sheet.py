import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

guro_file = '구로구 사용승인 현황(2020년이후).xlsx'

print("1. 구로구 엑셀 파일 읽기 시작...")
try:
    df = pd.read_excel(guro_file, sheet_name='사용승인_통합')
except Exception as e:
    print(f"오류: {guro_file} 파일을 읽지 못했습니다. {e}")
    sys.exit(1)

# 연면적이 숫자가 아닐 수 있으므로 numeric 변환 (오류 시 NaN)
df['연면적(㎡)'] = pd.to_numeric(df['연면적(㎡)'], errors='coerce')

print("2. 연면적 1만 이상 3만 미만 필터링...")
filtered = df[(df['연면적(㎡)'] >= 10000) & (df['연면적(㎡)'] < 30000)].copy().reset_index(drop=True)

print(f"  -> {len(filtered)}개의 건물이 추출되었습니다.")

print("3. 양천구 양식에 맞추어 데이터프레임 구성...")
new_df = pd.DataFrame()
new_df['연번'] = range(1, len(filtered) + 1)
new_df['건물명'] = filtered['건물명']
new_df['대지위치'] = filtered['대지위치']
new_df['연면적(제곱미터)'] = filtered['연면적(㎡)']
new_df['주용도'] = filtered['주용도']
new_df['건물관리업체 여부'] = ''
new_df['관리업체명'] = ''
new_df['관리업체 연락처'] = ''
new_df['허가일'] = filtered['허가일']
new_df['사용승인일'] = filtered['사용승인일']
new_df['최대지상층수'] = filtered['최대지상층수']
new_df['설계사무소명'] = filtered['설계사무소명']

print("4. 새로운 시트 '중대형건물_관리현황' 저장 중...")
try:
    with pd.ExcelWriter(guro_file, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        new_df.to_excel(writer, sheet_name='중대형건물_관리현황', index=False)
    print("구로구 파일에 '중대형건물_관리현황' 시트 생성 완료 (버전 1.0.0)")
except Exception as e:
    print(f"저장 중 오류 발생: {e}")
    sys.exit(1)
