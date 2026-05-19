import requests
from bs4 import BeautifulSoup
import sys
import ssl
from urllib3.util import create_urllib3_context

sys.stdout.reconfigure(encoding='utf-8')

# OpenSSL 보안 레벨을 낮추는 어댑터 정의
class CustomSSLAdapter(requests.adapters.HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        # SECLEVEL=1로 설정하여 WRONG_SIGNATURE_TYPE 오류 우회
        context = create_urllib3_context()
        context.check_hostname = False
        context.set_ciphers('DEFAULT@SECLEVEL=1')

        kwargs['ssl_context'] = context
        return super().init_poolmanager(*args, **kwargs)

url = 'https://ictis.kica.or.kr/construct/assessment/compSearch'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

session = requests.Session()
session.mount('https://', CustomSSLAdapter())

try:
    res = session.get(url, headers=headers, timeout=10, verify=False)
    print(f"Status Code: {res.status_code}")
    soup = BeautifulSoup(res.text, 'html.parser')
    
    # KICA의 검색 Ajax API 힌트 찾기
    scripts = soup.find_all('script')
    print(f"Found {len(scripts)} script tags")
    for i, script in enumerate(scripts):
        src = script.get('src', '')
        if src:
            print(f"  Script {i} src: {src}")
            
except Exception as e:
    print(f"Error: {e}")
