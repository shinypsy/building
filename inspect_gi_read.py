import sys
import requests
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding='utf-8')

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
}

url = "https://www.jobkorea.co.kr/Recruit/GI_Read/49186599?Oem_Code=C1"
print(f"Fetching GI_Read: {url}")

try:
    response = requests.get(url, headers=headers, timeout=10)
    print("Status code:", response.status_code)
    
    with open("gi_read.html", "w", encoding="utf-8") as f:
        f.write(response.text)
        
    soup = BeautifulSoup(response.text, 'html.parser')
    
    print("Title of GI_Read:", soup.title.string if soup.title else "No title")
    
    # Search for CEO, 매출액, 주소, 대표자 in text
    print("\nLooking for keywords...")
    for keyword in ["대표자", "매출액", "주소", "대표인", "자본금"]:
        found = []
        for el in soup.find_all(text=True):
            if keyword in el:
                found.append(el.strip())
        print(f"Keyword '{keyword}' matches: {len(found)}")
        for f in found[:3]:
            print("  -", f)
            
except Exception as e:
    print("Error:", e)
