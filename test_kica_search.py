import requests
from bs4 import BeautifulSoup
import sys
from urllib3.util import create_urllib3_context

sys.stdout.reconfigure(encoding='utf-8')

class CustomSSLAdapter(requests.adapters.HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        context = create_urllib3_context()
        context.check_hostname = False
        context.set_ciphers('DEFAULT@SECLEVEL=1')
        kwargs['ssl_context'] = context
        return super().init_poolmanager(*args, **kwargs)

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, http://g.co/chrome) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7'
}

session = requests.Session()
session.mount('https://', CustomSSLAdapter())

# 1. GET 방식으로 검색 시도
get_url = 'https://ictis.kica.or.kr/construct/compList'
params = {
    'searchSido': '',
    'searchType': '1',  # 업체명 검색 유형일 것
    'searchText': '계룡건설산업',
    'size': '10',
    'pageNumber': '1'
}

try:
    print("GET 요청으로 '계룡건설산업' 검색 시도 중...")
    res = session.get(get_url, headers=headers, params=params, timeout=10, verify=False)
    print(f"GET Status: {res.status_code}")
    print(f"Content Length: {len(res.text)}")
    
    # 결과 테이블 파싱
    soup = BeautifulSoup(res.text, 'html.parser')
    tables = soup.find_all('table')
    print(f"Found {len(tables)} tables in GET response")
    
    # 텍스트에 계룡건설산업이 들어 있는지 체크
    print(f"Raw Response Content: {res.text}")
    if '계룡건설산업' in res.text:
        print("-> 성공! GET 응답에 '계룡건설산업' 단어 포함됨!")

        # 첫 번째 테이블 내용 출력
        if tables:
            for row in tables[0].find_all('tr')[:5]:
                cols = [c.get_text(strip=True) for c in row.find_all(['th', 'td'])]
                print(f"  Row: {cols}")
    else:
        print("-> 실패. GET 응답에 계룡건설산업 없음.")
        
except Exception as e:
    print(f"GET Error: {e}")

# 2. POST 방식으로도 시도
try:
    print("\nPOST 요청으로 '계룡건설산업' 검색 시도 중...")
    res_post = session.post(get_url, headers=headers, data=params, timeout=10, verify=False)
    print(f"POST Status: {res_post.status_code}")
    print(f"Content Length: {len(res_post.text)}")
    
    if '계룡건설산업' in res_post.text:
        print("-> 성공! POST 응답에 '계룡건설산업' 단어 포함됨!")
        soup_post = BeautifulSoup(res_post.text, 'html.parser')
        tables_post = soup_post.find_all('table')
        if tables_post:
            for row in tables_post[0].find_all('tr')[:5]:
                cols = [c.get_text(strip=True) for c in row.find_all(['th', 'td'])]
                print(f"  Row: {cols}")
    else:
        print("-> 실패. POST 응답에 계룡건설산업 없음.")
except Exception as e:
    print(f"POST Error: {e}")
