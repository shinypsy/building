import requests
import re
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

# 카카오맵 웹 비공식 주소/장소 검색 API
url = 'https://search.map.kakao.com/mapsearch/map.daum'

params = {
    'q': '서울특별시 양천구 신월로36길 7',
    'msFlag': 'A',
    'sort': 'Accuracy',
    'output': 'json'
}


headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Referer': 'https://map.kakao.com/'
}

try:
    print("카카오맵 비공식 API로 '신정동 1025-28' 검색 중...")
    res = requests.get(url, headers=headers, params=params, timeout=5)
    print(f"Status Code: {res.status_code}")
    
    # 응답 텍스트가 JSON인지 확인하고 파싱
    data = res.json()
    print("Successfully parsed Kakao Map JSON response!")
    print(json.dumps(data, ensure_ascii=False, indent=2)[:3500])

        
except Exception as e:
    print(f"Error: {e}")
