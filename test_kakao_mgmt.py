import requests
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

url = 'https://search.map.kakao.com/mapsearch/map.daum'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
    'Referer': 'https://map.kakao.com/'
}

def search_mgmt_kakao(query):
    params = {
        'q': query,
        'msFlag': 'A',
        'sort': 'Accuracy',
        'output': 'json'
    }
    try:
        res = requests.get(url, headers=headers, params=params, timeout=5)
        if res.status_code == 200:
            data = res.json()
            places = data.get('place', [])
            if places:
                first = places[0]
                name = first.get('name', '')
                tel = first.get('tel', '')
                print(f"[{query}] -> {name} / {tel}")
                return name, tel
            else:
                print(f"[{query}] -> No results")
        else:
            print(f"API Error {res.status_code}")
    except Exception as e:
        print(f"Error: {e}")
    return None, None

search_mgmt_kakao('구로 신도림 팰러티움 관리사무소')
search_mgmt_kakao('오류동 삼전솔하임 관리사무소')
search_mgmt_kakao('구로동 186-7 예성유토피아 관리사무소')
search_mgmt_kakao('구로 칸타빌레 8차 관리사무소')
search_mgmt_kakao('구로미래에코타워 관리사무소')
search_mgmt_kakao('에드가 개봉 관리사무소')
search_mgmt_kakao('우분투 H포레스트 관리사무소')
