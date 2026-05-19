import requests
import json
import sys
import re
import time

sys.stdout.reconfigure(encoding='utf-8')

def get_building_name(address_str: str) -> str:
    # 1. 지번 주소 전처리 (외1필지, 외2필지 등 정제)
    clean_addr = re.sub(r'\s+외\d+필지.*', '', address_str).strip()
    
    # --- [Step 1] 카카오맵 비공식 API 검색 ---
    kakao_url = 'https://search.map.kakao.com/mapsearch/map.daum'
    kakao_params = {
        'q': clean_addr,
        'msFlag': 'A',
        'sort': 'Accuracy',
        'output': 'json'
    }
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Referer': 'https://map.kakao.com/'
    }
    
    try:
        res = requests.get(kakao_url, headers=headers, params=kakao_params, timeout=5)
        if res.status_code == 200:
            data = res.json()
            
            # (1) 주소 항목의 건물명 우선 체크
            addr_list = data.get('address', [])
            if addr_list and isinstance(addr_list, list):
                building_name = addr_list[0].get('building_name', '').strip()
                if building_name:
                    return building_name
                
                # 도로명주소 건물명 체크
                rel_building = addr_list[0].get('related_address_building_name', '').strip()
                if rel_building:
                    return rel_building
            
            # (2) 플레이스 항목의 상호명/장소명 체크
            place_list = data.get('place', [])
            if place_list and isinstance(place_list, list):
                place_name = place_list[0].get('name', '').strip()
                if place_name:
                    return place_name
    except Exception as e:
        pass
        
    # --- [Step 2] 도로명주소 공식 API 폴백 검색 ---
    juso_url = 'https://www.juso.go.kr/addrlink/addrLinkApiJsonp.do'
    juso_params = {
        'confmKey': 'U01TX0FVVEhSMjAxODEwMjUxNTAzMTAxMDgyNTM=',
        'keyword': clean_addr,
        'resultType': 'json',
        'currentPage': '1',
        'countPerPage': '5'
    }
    try:
        res = requests.get(juso_url, headers=headers, params=juso_params, timeout=5)
        if res.status_code == 200:
            text = res.text
            start_idx = text.find('(')
            end_idx = text.rfind(')')
            if start_idx != -1 and end_idx != -1:
                json_str = text[start_idx + 1:end_idx]
                data = json.loads(json_str)
                juso_list = data.get('results', {}).get('juso', [])
                if juso_list:
                    bd_nm = juso_list[0].get('bdNm', '').strip()
                    if bd_nm:
                        return bd_nm
    except Exception as e:
        pass
        
    # 모든 수단이 실패했을 경우 공백 리턴
    return ""

# 샘플 주소 테스트 구동
test_addresses = [
    '서울특별시 양천구 신정동 1025-28',
    '서울특별시 양천구 신정동 1190-19',
    '서울특별시 양천구 신정동 880-6 외1필지',
    '서울특별시 양천구 목동 547-4',
    '서울특별시 양천구 목동 404-114 외1필지',
    '서울특별시 양천구 목동 324-26',
    '서울특별시 양천구 신정동 988-3 외1필지',
    '서울특별시 양천구 목동 404-156',
    '서울특별시 양천구 신정동 118-1'
]

print("3중 하이브리드 건물명 매칭 엔진 가동 테스트...")
for addr in test_addresses:
    t_start = time.time()
    bname = get_building_name(addr)
    t_end = time.time()
    print(f"주소: {addr:40s} -> 건물명: {bname:25s} (소요시간: {t_end - t_start:.2f}초)")
    time.sleep(0.3)
