import requests
from bs4 import BeautifulSoup
import sys
import ssl
from urllib3.util import create_urllib3_context

sys.stdout.reconfigure(encoding='utf-8')

class CustomSSLAdapter(requests.adapters.HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
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
    soup = BeautifulSoup(res.text, 'html.parser')
    
    # 1. form 태그 검출
    forms = soup.find_all('form')
    print(f"Found {len(forms)} form tags:")
    for f in forms:
        print(f"  Form: id={f.get('id')}, name={f.get('name')}, action={f.get('action')}, method={f.get('method')}")
        
    # 2. input 태그 중 text 타입 검출
    inputs = soup.find_all('input')
    print(f"\nFound {len(inputs)} input tags:")
    for i in inputs:
        print(f"  Input: id={i.get('id')}, name={i.get('name')}, type={i.get('type')}, class={i.get('class')}")
        
    # 3. button 태그 검출
    buttons = soup.find_all(['button', 'a'])
    print(f"\nFound button/link tags:")
    for b in buttons:
        b_id = b.get('id')
        b_class = b.get('class')
        b_text = b.get_text(strip=True)
        if b_id or 'btn' in str(b_class) or '검색' in b_text:
            print(f"  Button/Link: id={b_id}, class={b_class}, text={b_text}")
            
except Exception as e:
    print(f"Error: {e}")
