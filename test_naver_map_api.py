import requests
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

url = 'https://map.naver.com/v5/api/search'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
    'Referer': 'https://map.naver.com/'
}

params = {
    'caller': 'pcweb',
    'query': '서울특별시 양천구 신정동 1025-28',
    'type': 'all'
}

try:
    print("네이버 지도 API로 '신정동 1025-28' 검색 중...")
    res = requests.get(url, headers=headers, params=params, timeout=5)
    print(f"Status Code: {res.status_code}")
    
    data = res.json()
    print("Successfully parsed Naver Map API JSON response!")
    print(json.dumps(data, ensure_ascii=False, indent=2)[:3000])
        
except Exception as e:
    print(f"Error: {e}")


