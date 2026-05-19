import pandas as pd
import requests
from urllib3.util import create_urllib3_context
import urllib3
import time
import random
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class CustomSSLAdapter(requests.adapters.HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        # OpenSSL 3.0 SECLEVEL=1 및 check_hostname=False 설정하여 SSL 서명 에러 우회
        context = create_urllib3_context()
        context.check_hostname = False
        context.set_ciphers('DEFAULT@SECLEVEL=1')
        kwargs['ssl_context'] = context
        return super().init_poolmanager(*args, **kwargs)

def verify_kica_registration() -> None:
    file_path = 'FM 업체 리스트_정보포함.xlsx'
    print(f"[{file_path}] 파일을 로드하여 KICA 공식 API 회원 조회를 시작합니다...")
    df = pd.read_excel(file_path)
    
    # 1. 기존 정보통신공사 관련 2개 항목 전격 제거
    cols_to_drop = ['정보통신공사업 여부', '정보통신공사면허 여부']
    for col in cols_to_drop:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)
            print(f" -> 기존 컬럼 '{col}'을 성공적으로 제거했습니다.")
            
    telecom_reg_list: list[str] = []
    
    # 2. KICA 연계 세션 설정
    session = requests.Session()
    session.mount('https://', CustomSSLAdapter())
    
    kica_url = 'https://ictis.kica.or.kr/construct/compList'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    
    total_rows = len(df)
    
    # 인풋 페이징(Input Paging) 구조: 각 행의 인풋 업체명을 기반으로 점진적 순회
    for idx, row in df.iterrows():
        name = str(row['업체명']).strip()
        
        # 주식회사, ㈜ 등 불필요한 회사 접두/접미사 정제
        clean_name = re.sub(r'\(주\)|주식회사|㈜', '', name).strip()
        
        print(f"[{idx+1}/{total_rows}] '{name}' -> KICA 실시간 조회 중...", end='', flush=True)
        
        params = {
            'searchSido': '',
            'searchType': '1',  # 업체명 검색
            'searchText': clean_name,
            'size': '10',
            'pageNumber': '1'
        }
        
        is_registered = "N"
        
        try:
            res = session.get(kica_url, headers=headers, params=params, timeout=10, verify=False)
            if res.status_code == 200:
                res_data = res.json()
                kica_list = res_data.get('data', {}).get('list', [])
                
                # 결과 매칭 판정
                if kica_list:
                    # 첫 번째 회원사의 한글명 가져와서 유사도 체크
                    firm_name = kica_list[0].get('firmNmKor', '')
                    firm_clean = re.sub(r'\(주\)|주식회사|㈜', '', firm_name).strip()
                    
                    # 회사명 상호 포함 관계 매칭 체크
                    if clean_name in firm_clean or firm_clean in clean_name:
                        is_registered = "Y"
                        print(f" -> 등록 확인 (KICA 등록상호: {firm_name})")
                    else:
                        print(f" -> 불일치 (KICA 등록상호: {firm_name} / 대조군: {clean_name})")
                else:
                    print(" -> 회원사 없음")
            else:
                print(f" -> HTTP {res.status_code} 에러")
        except Exception as e:
            print(f" -> 조회 에러 ({e})")
            
        telecom_reg_list.append(is_registered)
        time.sleep(random.uniform(0.3, 0.7))
        
    df['정보통신공사업 등록'] = telecom_reg_list
    
    # 덮어쓰기
    try:
        df.to_excel(file_path, index=False)
        print(f"\n모든 KICA API 조회 및 매칭이 완료되어 '{file_path}' 파일에 최종 반영되었습니다!")
    except PermissionError:
        backup_path = 'FM 업체 리스트_정보포함_v2.xlsx'
        df.to_excel(backup_path, index=False)
        print(f"\n[알림] '{file_path}' 파일이 현재 열려 있어 '{backup_path}'로 우회 저장되었습니다!")

if __name__ == '__main__':
    verify_kica_registration()
