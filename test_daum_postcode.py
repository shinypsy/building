import requests
import re
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Daum 우편번호 서비스 공개용 검색 API
url = 'https://postcode.map.daum.net/search'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Referer': 'https://postcode.map.daum.net/'
}

params = {
    'q': '서울특별시 양천구 신정동 1025-28',
    'origin': 'https://postcode.map.daum.net'
}

try:
    print("Daum 우편번호 API로 '신정동 1025-28' 검색 중...")
    res = requests.get(url, headers=headers, params=params, timeout=5)
    print(f"Status Code: {res.status_code}")
    
    # Daum 우편번호 API는 JSONP 형식이거나 특수한 문자열 형태일 수 있으므로 정규식으로 파싱
    text = res.text
    print(f"Raw Content Length: {len(text)}")
    
    # 응답 텍스트 일부 출력
    print(f"Sample Text:\n{text[:1000]}")
    
    # 건물명(buildingName) 단어 패턴 탐색
    if '신정그린빌라' in text:
        print("-> 성공! Daum 우편번호 API 응답에 '신정그린빌라' 단어 존재함!")
        
except Exception as e:
    print(f"Error: {e}")
