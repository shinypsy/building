import requests
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

# PC 네이버 지도 주소 전용 검색 API
url = 'https://map.naver.com/v5/api/addresses'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'application/json, text/plain, */*',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
    'Referer': 'https://map.naver.com/'
}

params = {
    'query': '서울특별시 양천구 신정동 1025-28'
}

try:
    print("네이버 지도 주소 API로 '신정동 1025-28' 검색 중...")
    res = requests.get(url, headers=headers, params=params, timeout=5)
    print(f"Status Code: {res.status_code}")
    
    if res.status_code == 200:
        data = res.json()
        print("Success to fetch address data!")
        # JSON 예쁘게 출력
        print(json.dumps(data, ensure_ascii=False, indent=2)[:2000])
    else:
        print(f"Failed. Response: {res.text[:500]}")
        
except Exception as e:
    print(f"Error: {e}")
