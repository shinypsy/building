import requests
from bs4 import BeautifulSoup
import sys
import re

sys.stdout.reconfigure(encoding='utf-8')

url = 'https://www.google.com/search'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}
params = {
    'q': '서울특별시 양천구 신정동 1025-28',
    'hl': 'ko'  # 한글 검색 결과 우선
}

try:
    print("구글 검색으로 '신정동 1025-28' 검색 중...")
    res = requests.get(url, headers=headers, params=params, timeout=5)
    print(f"Status Code: {res.status_code}")
    
    soup = BeautifulSoup(res.text, 'html.parser')
    
    # 전체 텍스트에서 '신정그린빌라' 확인
    if '신정그린빌라' in res.text:
        print("-> 성공! 구글 검색 HTML 내에 '신정그린빌라' 텍스트 존재함!")
        
    # 검색 결과 제목(h3 태그) 또는 스니펫(span 이나 div) 출력해보기
    h3_tags = soup.find_all('h3')
    print(f"\nFound {len(h3_tags)} h3 tags in search results:")
    for h3 in h3_tags[:5]:
        print(f"  H3 Title: {h3.get_text(strip=True)}")
        
    # 검색 스니펫 영역 텍스트 추출
    print("\n[Snippet Texts]")
    snippets = soup.find_all('span', class_=re.compile(r'aCOn1c|HG397|VwiC3b'))
    for snip in snippets[:5]:
        print(f"  Snippet: {snip.get_text(strip=True)}")
        
except Exception as e:
    print(f"Error: {e}")
