import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import time
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

EXCEL_PATH = Path("guro_use_2026.xlsx")
# 일반적인 한국 전화번호 및 서울 유선전화/대표전화 탐색 정규식 (하이픈 포함 혹은 미포함)
PHONE_REGEX = re.compile(r"\b(0\d{1,2}[-]?\d{3,4}[-]?\d{4})\b")
HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Referer': 'https://search.naver.com/'
}

def search_phone_number(query):
    url = 'https://search.naver.com/search.naver'
    params = {'query': query}
    try:
        res = requests.get(url, headers=HEADERS, params=params, timeout=5)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            # HTML 태그를 제거하고 텍스트만 추출하여 쪼개진 전화번호 합침
            clean_text = soup.get_text(separator=' ')
            matches = PHONE_REGEX.findall(clean_text)
            if matches:
                # 중복 제거
                unique_matches = list(dict.fromkeys(matches))
                # 010 번호는 가급적 제외하고 유선전화 우선 선택
                for num in unique_matches:
                    if not num.startswith("010"):
                        return num
                return unique_matches[0]
    except Exception as e:
        print(f"검색 오류 ({query}): {e}")
    return None

def update_contacts_with_naver(start_row=0, batch_size=25):
    """
    인풋 페이징(Input Paging)을 지원하여, 지정한 범위 내에서 네이버 검색을 진행해 엑셀의 new tel 컬럼을 업데이트한다.
    """
    print(f"=== 네이버 검색 기반 전화번호 업데이트 개시 (범위: {start_row} ~ {start_row + batch_size - 1}) ===")
    
    # 엑셀 시트 읽기
    xls = pd.ExcelFile(EXCEL_PATH)
    sheets = {name: pd.read_excel(xls, name) for name in xls.sheet_names}
    df = sheets['중대형건물_관리현황']
    
    # 'new tel' 열이 없는 경우 생성
    if 'new tel' not in df.columns:
        df['new tel'] = '-'
        
    total_rows = len(df)
    end_row = min(start_row + batch_size, total_rows)
    
    updated_count = 0
    for idx in range(start_row, end_row):
        row = df.iloc[idx]
        bname = str(row.get('건물명', '')).strip()
        manager = str(row.get('관리업체명', '')).strip()
        addr = str(row.get('대지위치', '')).strip()
        
        # 1. 쿼리 생성 (구/동 정보 + 관리업체명 조합)
        addr_match = re.search(r'(구로구\s+\S+동)', addr)
        region = addr_match.group(1) if addr_match else '구로구'
        
        query = f"{region} {manager}"
        print(f"\n[{idx + 1}/{total_rows}] '{bname}' 검색 진행 중 -> 쿼리: {query}")
        
        phone = search_phone_number(query)
        
        # 2차 Fallback: 건물명 + 관리사무소로 재검색
        if not phone and bname != 'nan':
            fallback_query = f"{region} {bname} 관리사무소"
            print(f"  └ 1차 실패, Fallback 검색 진행 중 -> 쿼리: {fallback_query}")
            phone = search_phone_number(fallback_query)
            
        if phone:
            print(f"  ✔ 전화번호 획득 성공: {phone}")
            df.at[idx, 'new tel'] = phone
            updated_count += 1
        else:
            print(f"  ❌ 전화번호 검색 실패")
            df.at[idx, 'new tel'] = '-'
            
        time.sleep(1.5)  # IP 차단 방지 딜레이
        
    # 시트 복구 후 덮어쓰기 저장
    sheets['중대형건물_관리현황'] = df
    
    saved = False
    retry_count = 0
    while not saved and retry_count < 3:
        try:
            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
                for name, sheet_df in sheets.items():
                    sheet_df.to_excel(writer, sheet_name=name, index=False)
            saved = True
            print(f"\n=== 업데이트 프로세스 완료! 총 {updated_count}건의 연락처가 new tel 열에 적재 및 저장되었습니다. ===")
        except PermissionError:
            retry_count += 1
            print(f"\n[경고] '{EXCEL_PATH}' 파일이 현재 다른 프로그램에 의해 열려 있어 쓰기 권한이 없습니다.")
            if retry_count < 3:
                print("5초 후에 저장을 재시도합니다. 그동안 실행 중인 엑셀 프로그램을 닫아주세요...")
                time.sleep(5)
            else:
                fallback_path = EXCEL_PATH.parent / f"{EXCEL_PATH.stem}_업데이트_임시{EXCEL_PATH.suffix}"
                print(f"재시도 실패! 데이터를 보존하기 위해 대체 파일 '{fallback_path.name}'로 임시 저장합니다.")
                with pd.ExcelWriter(fallback_path, engine='openpyxl') as writer:
                    for name, sheet_df in sheets.items():
                        sheet_df.to_excel(writer, sheet_name=name, index=False)
                saved = True
                
if __name__ == '__main__':
    # 인풋 페이징 방식으로 0행부터 25행까지 전수 안전 처리
    update_contacts_with_naver(start_row=0, batch_size=25)
