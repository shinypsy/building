import pandas as pd
import requests
import json
import sys
import re
import time
import random
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

sys.stdout.reconfigure(encoding='utf-8')

# 메모리 캐시 및 멀티스레드 락 설정
address_cache = {}
cache_lock = threading.Lock()
print_lock = threading.Lock()

def get_building_name(address_str: str, session: requests.Session) -> str:
    # 1. 대지위치 지번 주소 정제 (외1필지 등 행정 텍스트 제거)
    clean_addr = re.sub(r'\s+외\d+필지.*', '', address_str).strip()
    if not clean_addr:
        return ""
    
    # 스레드 세이프 캐시 확인
    with cache_lock:
        if clean_addr in address_cache:
            return address_cache[clean_addr]
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Referer': 'https://map.kakao.com/'
    }
    
    # --- [Step 1 & 2] 카카오맵 비공식 API 검색 ---
    kakao_url = 'https://search.map.kakao.com/mapsearch/map.daum'
    kakao_params = {
        'q': clean_addr,
        'msFlag': 'A',
        'sort': 'Accuracy',
        'output': 'json'
    }
    
    try:
        res = session.get(kakao_url, headers=headers, params=kakao_params, timeout=5)
        if res.status_code == 200:
            data = res.json()
            
            # 주소 항목의 건물명 우선 체크
            addr_list = data.get('address', [])
            if addr_list and isinstance(addr_list, list):
                building_name = addr_list[0].get('building_name', '').strip()
                if building_name:
                    with cache_lock:
                        address_cache[clean_addr] = building_name
                    return building_name
                
                # 연계 도로명 건물명 체크
                rel_building = addr_list[0].get('related_address_building_name', '').strip()
                if rel_building:
                    with cache_lock:
                        address_cache[clean_addr] = rel_building
                    return rel_building
            
            # 플레이스(상호/장소명) 체크
            place_list = data.get('place', [])
            if place_list and isinstance(place_list, list):
                place_name = place_list[0].get('name', '').strip()
                if place_name:
                    with cache_lock:
                        address_cache[clean_addr] = place_name
                    return place_name
    except Exception:
        pass
        
    # --- [Step 3] 도로명주소 오픈 API 폴백 검색 ---
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
            start_idx = text.find('(')
            end_idx = text.rfind(')')
            if start_idx != -1 and end_idx != -1:
                json_str = text[start_idx + 1:end_idx]
                json_data = json.loads(json_str)
                juso_list = json_data.get('results', {}).get('juso', [])
                if juso_list:
                    bd_nm = juso_list[0].get('bdNm', '').strip()
                    if bd_nm:
                        with cache_lock:
                            address_cache[clean_addr] = bd_nm
                        return bd_nm
    except Exception:
        pass
        
    # 매칭 실패 캐싱
    with cache_lock:
        address_cache[clean_addr] = ""
    return ""

def process_single_row(idx, addr, session):
    # 단일 주소에 대한 건물명 매칭 및 미세 랜덤 지연
    bname = get_building_name(addr, session)
    time.sleep(random.uniform(0.05, 0.15))
    return idx, bname

def process_yangcheon_buildings():
    file_path = '양천구_건축허가 및 사용승인현황_20260311.csv'
    print(f"[{file_path}] 파일을 로딩합니다...", flush=True)
    
    try:
        df = pd.read_csv(file_path, encoding='cp949')
    except Exception as e:
        print(f"로드 실패: {e}", flush=True)
        return
        
    # '건물명' 컬럼이 없는 경우 생성
    if '건물명' not in df.columns:
        df['건물명'] = ""
        
    # 미완료 대상 선별
    df['건물명'] = df['건물명'].fillna("").astype(str)
    tasks = []
    
    for idx, row in df.iterrows():
        current_bname = row['건물명'].strip()
        if not current_bname or current_bname == 'nan':
            tasks.append((idx, str(row['대지위치']).strip()))
            
    total_tasks = len(tasks)
    print(f"총 {len(df)}개 행 중 미처리된 {total_tasks}개 주소지에 대해 병렬 매칭을 가동합니다...", flush=True)
    
    if total_tasks == 0:
        print("모든 주소지의 건물명이 이미 매칭 완료되었습니다!", flush=True)
        return
        
    # 스레드 세션 설정
    session = requests.Session()
    success_count = len(df) - total_tasks
    
    # 20개의 고속 병렬 스레드 풀 구동
    with ThreadPoolExecutor(max_workers=20) as executor:
        futures = {executor.submit(process_single_row, idx, addr, session): idx for idx, addr in tasks}
        
        for i, future in enumerate(as_completed(futures)):
            try:
                idx, bname = future.result()
                df.at[idx, '건물명'] = bname
                if bname:
                    success_count += 1
            except Exception as e:
                pass
                
            # 실시간 진행 출력 (Unbuffered)
            if (i + 1) % 50 == 0 or i + 1 == total_tasks:
                with print_lock:
                    print(f" -> 병렬 매칭 진행률: [{i + 1}/{total_tasks}] 완료 (총 확보 건물명: {success_count}건)", flush=True)
                    
            # 200행 단위 실시간 덮어쓰기 저장 (중간 유실 원천 방지)
            if (i + 1) % 200 == 0:
                with cache_lock:
                    df.to_csv(file_path, index=False, encoding='cp949')
                with print_lock:
                    print(f"   [체크포인트] {i+1}개 행 가공 결과 파일 세이브 완료.", flush=True)
                    
    # 최종 결과 덮어쓰기
    try:
        df.to_csv(file_path, index=False, encoding='cp949')
        print(f"\n모든 병렬 수집이 성황리에 끝났습니다! 최종본이 '{file_path}'에 저장되었습니다.", flush=True)
        print(f"총 {len(df)}건 중 {success_count}건의 실시간 건물명 획득 완료!", flush=True)
    except Exception as e:
        print(f"최종 저장 중 오류 발생: {e}", flush=True)

if __name__ == '__main__':
    process_yangcheon_buildings()
