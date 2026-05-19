import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel('경비경호 업체 순위_정보포함.xlsx')
print(df[['순위', '업체명', '매출액(원)', '매출액 기준일', '사업장 주소']].head(20))
