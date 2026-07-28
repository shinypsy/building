import requests
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

url = 'https://m.map.naver.com/api/place/v5/search'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Referer': 'https://m.map.naver.com/'
}

def search_mgmt(query):
    params = {
        'query': query,
        'page': '1'
    }
    try:
        res = requests.get(url, headers=headers, params=params, timeout=5)
        if res.status_code == 200:
            data = res.json()
            items = data.get('result', {}).get('place', {}).get('list', [])
            if items:
                first = items[0]
                name = first.get('name', '')
                tel = first.get('tel', '')
                addr = first.get('roadAddress', '') or first.get('address', '')
                print(f"[{query}] -> {name} / {tel} / {addr}")
            else:
                print(f"[{query}] -> No results")
        else:
            print(f"API Error {res.status_code}")
    except Exception as e:
        print(f"Error: {e}")

search_mgmt('구로 신도림 팰러티움 관리사무소')
search_mgmt('오류동 삼전솔하임 관리사무소')
search_mgmt('구로 에드가 개봉 관리사무소')
search_mgmt('구로 칸타빌레 8차 관리사무소')
search_mgmt('고척동 76-160 관리사무소')
search_mgmt('개봉동 170-18 관리사무소')
