import requests
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

# 모바일 네이버 지도 플레이스 검색 API
url = 'https://m.map.naver.com/api/place/v5/search'
headers = {
    'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1',
    'Referer': 'https://m.map.naver.com/'
}

params = {
    'query': '서울특별시 양천구 신정동 1025-28',
    'requester': 'html5',
    'page': '1'
}

try:
    print("네이버 지도 플레이스 API로 검색 중...")
    res = requests.get(url, headers=headers, params=params, timeout=5)
    print(f"Status Code: {res.status_code}")
    
    if res.status_code == 200:
        data = res.json()
        print("Success!")
        # JSON 예쁘게 출력
        print(json.dumps(data, ensure_ascii=False, indent=2)[:3500])
    else:
        print(f"Failed. Response: {res.text[:500]}")
        
except Exception as e:
    print(f"Error: {e}")
