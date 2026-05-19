import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import random
import sys

sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel('FM 업체 리스트_정보포함.xlsx')
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

print("상위 15개 업체에 대해 '정보통신' 및 '정보통신공사' 키워드 매칭 테스트를 진행합니다...")
for i, row in df.head(15).iterrows():
    name = row['업체명']
    link = row['잡코리아 링크']
    
    if pd.isna(link) or link == 'N/A':
        print(f"[{i+1}] {name}: 잡코리아 링크 없음 (N/A)")
        continue
        
    try:
        res = requests.get(link, headers=headers, timeout=5)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            text = soup.get_text()
            
            # 정보통신 및 정보통신공사 여부 체크
            has_telecom = '정보통신' in text
            has_telecom_biz = '정보통신공사' in text
            
            # 구체적인 문구 확인을 위해 매칭 문구들 추출
            matched_sentences = []
            for line in text.split('\n'):
                line = line.strip()
                if '정보통신' in line:
                    matched_sentences.append(line[:100])
                    
            print(f"[{i+1}] {name}: 정보통신={has_telecom}, 정보통신공사={has_telecom_biz}")
            if matched_sentences:
                print(f"    매칭 문장 예시: {matched_sentences[:2]}")
        else:
            print(f"[{i+1}] {name}: HTTP {res.status_code}")
    except Exception as e:
        print(f"[{i+1}] {name}: 에러 발생 {e}")
        
    time.sleep(random.uniform(0.3, 0.7))
