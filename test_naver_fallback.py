import requests
import urllib.parse
from bs4 import BeautifulSoup
import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
}

def search_naver_fallback(company_name):
    print(f"NAVER Fallback search for '{company_name}'...")
    query = urllib.parse.quote(f"{company_name} 잡코리아")
    url = f"https://search.naver.com/search.naver?query={query}"
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Look for any links containing jobkorea.co.kr/Recruit/Co_Read
        co_read_links = []
        for a in soup.find_all('a', href=True):
            href = a['href']
            if 'jobkorea.co.kr' in href and ('Co_Read' in href or 'corp' in href):
                co_read_links.append(href)
                
        print(f"Found {len(co_read_links)} links:")
        for l in co_read_links:
            print("  -", l)
            m = re.search(r'/Co_Read/C/(\d+)', l)
            if m:
                print("  Extracted ID:", m.group(1))
                return m.group(1)
        return None
    except Exception as e:
        print("NAVER fallback error:", e)
        return None

search_naver_fallback("(주)맥서브")
