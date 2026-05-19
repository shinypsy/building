import sys
import requests
from bs4 import BeautifulSoup
import urllib.parse
import json

sys.stdout.reconfigure(encoding='utf-8')

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
}

def test_search(company_name):
    query = urllib.parse.quote(company_name)
    url = f"https://www.jobkorea.co.kr/Search/?stext={query}"
    print(f"Searching: {url}")
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        print("Status code:", response.status_code)
        
        # Save html to examine if needed
        with open("search_result.html", "w", encoding="utf-8") as f:
            f.write(response.text)
            
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Let's find any links containing /corp/ or company profile
        links = []
        for a in soup.find_all('a', href=True):
            href = a['href']
            text = a.get_text(strip=True)
            if '/corp/' in href or 'company' in href or 'co.kr/Recruit/Co_Read' in href:
                links.append({'href': href, 'text': text})
                
        print(f"Found {len(links)} links related to corporate info:")
        for l in links[:10]:
            print(l)
            
    except Exception as e:
        print("Error:", e)

test_search("(주)에스앤아이코퍼레이션")
