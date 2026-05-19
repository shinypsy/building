import requests
from bs4 import BeautifulSoup
import sys
import re

sys.stdout.reconfigure(encoding='utf-8')

url = 'https://search.naver.com/search.naver'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}
params = {
    'query': '서울특별시 양천구 신정동 1025-28'
}

try:
    print("네이버 통합검색으로 '신정동 1025-28' 검색 중...")
    res = requests.get(url, headers=headers, params=params, timeout=5)
    print(f"Status Code: {res.status_code}")
    
    soup = BeautifulSoup(res.text, 'html.parser')
    
    # 1. 텍스트 내에서 신정그린빌라 검색
    if '신정그린빌라' in res.text:
        print("-> 성공! HTML 내에 '신정그린빌라' 텍스트 존재함!")
        
    # 2. 지도 영역이나 장소 관련 클래스명 검출
    # 네이버 통합검색의 주소 검색 결과 영역 파싱
    # 보통 class="api_subject_bx" 나 class="addr_title" 등의 클래스를 가짐
    print("\n[주소/장소 관련 태그 탐색]")
    addr_tags = soup.find_all(class_=re.compile(r'addr|title|place|name'))
    for tag in addr_tags[:15]:
        tag_text = tag.get_text(strip=True)
        if len(tag_text) > 1 and len(tag_text) < 50:
            print(f"  Class: {tag.get('class')}, Text: {tag_text}")
            
    # 전체 텍스트 중 주소 주변의 텍스트 확인
    print("\n[신정그린빌라 매칭 부분의 주변 HTML 출력]")
    for match in re.finditer(r'신정그린빌라', res.text):
        start = max(0, match.start() - 100)
        end = min(len(res.text), match.end() + 100)
        print(f"--- Match ---\n{res.text[start:end]}\n-------------")
        
except Exception as e:
    print(f"Error: {e}")
