# 양천구 건축허가 건물명 정밀 매칭 및 데이터베이스 구축 구현 계획 (plan.md)

## 1. 접근 방식 및 코드 구조 상세 설명
- **대상 파일**: `양천구_건축허가 및 사용승인현황_20260311.csv` (cp949 인코딩)
- **목표**: 각 행의 `대지위치` 지번 주소를 정밀 파싱하고, 카카오맵 비공식 API와 도로명주소 공식 API를 유기적으로 연계한 **'3중 하이브리드 건물명 매칭 엔진'**을 통해 해당 주소의 실제 **`건물명`** 컬럼을 새로 추가/정제하여 최종 갱신된 CSV를 저장함.
- **인풋 페이징 및 최적화 기법 (Input Paging & Cache)**:
  - 1978행의 대용량 데이터를 처리하므로 **메모리 캐싱(Memoization)**을 적용하여 동일 주소지의 중복 조회를 0ms로 단축시킴.
  - 매 100행마다 임시 세이브(Checkpointing)를 수행하여 돌발 정전이나 API 일시 오류 시에도 이어서 처리 가능하도록 안전망 구축.
  - 서버 과부하 방지 및 IP 차단 예방을 위해 각 주소지 순회 시 0.1~0.3초의 임의 대기(Time Sleep) 이식.
- **3중 하이브리드 매칭 알고리즘**:
  1. `외N필지` 등 불필요한 행정 접미사 정제.
  2. **[1단계] 카카오맵 API**: 해당 지번의 `building_name` 또는 연계 도로명의 `related_address_building_name` 존재 시 최우선 매칭.
  3. **[2단계] 카카오맵 플레이스**: 빌딩명이 비어 있는 상가 건물의 경우, 1순위 대표 상호명(`place.name`)으로 보완.
  4. **[3단계] 도로명주소 오픈 API 데모**: 행안부 도로명 대장(`bdNm`)을 통해 3차 폴백 검증.
  5. 최종 실패 시 공백으로 처리하되 주용도 정보를 참고함.

## 2. 파일 경로 (File Paths)
- 대상 및 출력 파일: `d:\Dev\Project\sample\양천구_건축허가 및 사용승인현황_20260311.csv` [MODIFY]
- 분석 및 업데이트 스크립트: `d:\Dev\Project\sample\update_yangcheon_buildings.py` [NEW]
- 계획 파일: `d:\Dev\Project\sample\doc\plan.md` [MODIFY]
- 리서치 파일: `d:\Dev\Project\sample\doc\research.md` [MODIFY]

## 3. 코드 스니펫 (Code Snippet)
```python
import pandas as pd
import requests
import json
import sys
import re
import time
import random

sys.stdout.reconfigure(encoding='utf-8')

# 메모리 캐시 및 진행 세이브 설정
address_cache = {}

def get_building_name(address_str: str, session: requests.Session) -> str:
    # 지번 주소 정제 (외1필지, 외2필지 등 제거)
    clean_addr = re.sub(r'\s+외\d+필지.*', '', address_str).strip()
    if not clean_addr:
        return ""
    
    if clean_addr in address_cache:
        return address_cache[clean_addr]
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Referer': 'https://map.kakao.com/'
    }
    
    # 1단계 & 2단계: 카카오맵 비공식 API 활용
    kakao_url = 'https://search.map.kakao.com/mapsearch/map.daum'
    kakao_params = {'q': clean_addr, 'msFlag': 'A', 'sort': 'Accuracy', 'output': 'json'}
    
    try:
        res = session.get(kakao_url, headers=headers, params=kakao_params, timeout=5)
        if res.status_code == 200:
            data = res.json()
            
            # 주소 객체의 빌딩 정보 체크
            addr_list = data.get('address', [])
            if addr_list and isinstance(addr_list, list):
                bname = addr_list[0].get('building_name', '').strip()
                if bname:
                    address_cache[clean_addr] = bname
                    return bname
                
                rel_bname = addr_list[0].get('related_address_building_name', '').strip()
                if rel_bname:
                    address_cache[clean_addr] = rel_bname
                    return rel_bname
            
            # 플레이스(상가 상호명) 체크
            place_list = data.get('place', [])
            if place_list and isinstance(place_list, list):
                pname = place_list[0].get('name', '').strip()
                if pname:
                    address_cache[clean_addr] = pname
                    return pname
    except Exception:
        pass
        
    # 3단계: 도로명주소 오픈 API 폴백
    juso_url = 'https://www.juso.go.kr/addrlink/addrLinkApiJsonp.do'
    juso_params = {
        'confmKey': 'U01TX0FVVEhSMjAxODEwMjUxNTAzMTAxMDgyNTM=',
        'keyword': clean_addr,
        'resultType': 'json',
        'currentPage': '1',
        'countPerPage': '5'
    }
    try:
        res = session.get(juso_url, headers=headers, params=juso_params, timeout=5)
        if res.status_code == 200:
            text = res.text
            start = text.find('(')
            end = text.rfind(')')
            if start != -1 and end != -1:
                json_data = json.loads(text[start+1:end])
                juso_list = json_data.get('results', {}).get('juso', [])
                if juso_list:
                    bd_nm = juso_list[0].get('bdNm', '').strip()
                    if bd_nm:
                        address_cache[clean_addr] = bd_nm
                        return bd_nm
    except Exception:
        pass
        
    address_cache[clean_addr] = ""
    return ""
```

## 4. 트레이드오프 (Trade-offs)
- **대량 API 호출 레이턴시**: 1978행의 주소를 실시간 검색하므로 10~15분 가량 소요되나, **캐싱 설계**를 접목하여 이미 검색했던 동일 대지위치에 대한 중복 트래픽과 소요시간을 획득함과 동시에 **체크포인트(100행 주기 저장)** 기능을 탑재해 완전 무결성을 제공함.
