import pandas as pd

filepath = '구로구 사용승인 현황(2020년이후).xlsx'
print('파일 로딩중...')
df = pd.read_excel(filepath, sheet_name='사용승인_통합')

cols_to_numeric = ['대지면적(㎡)', '건축면적(㎡)', '연면적(㎡)', '증축연면적(㎡)', '총주차장면적(㎡)']
print('숫자 변환중...')

for col in cols_to_numeric:
    if col in df.columns:
        # 천 단위 쉼표 제거 후 float으로 강제 변환
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', '').str.strip(), errors='coerce')

with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='사용승인_통합')

print('성공적으로 숫자 변환 완료 및 저장되었습니다!')
