import os, pandas as pd
output_path = os.path.join(os.getcwd(), '구로구 건축허가 현황(2020년이후).xlsx')
# 샘플 데이터
data = {
    '허가일': ['2020-01-15','2021-06-20','2022-03-05'],
    '건물명': ['예시아파트','예시오피스텔','예시주상복합'],
    '연면적(제곱미터)': [1200,850,2500],
    '출처 URL': ['https://example.com/1','https://example.com/2','https://example.com/3']
}
df = pd.DataFrame(data)
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name='건축허가')
print('Created', output_path)
