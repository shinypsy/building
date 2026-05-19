import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel('시설관리 업체 순위_정보포함.xlsx')
failed_rows = df[df['업체명'].isin(['(주)미래에이비엠', '한국북부발전(주)', '파주도시공사'])]
print(failed_rows)
