import sys
import requests
import re
from bs4 import BeautifulSoup
import urllib.parse
import json

sys.stdout.reconfigure(encoding='utf-8')

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
}

def clean_company_name(name):
    # Remove (주), ㈜, (주 ), (유) etc.
    name = re.sub(r'[\(（]주[\)）]|㈜|[\(（]유[\)）]|[\(（]합[\)）]', '', name)
    return name.strip()

def search_jobkorea(company_name):
    print(f"\n--- Searching JobKorea for '{company_name}' ---")
    query = urllib.parse.quote(company_name)
    url = f"https://www.jobkorea.co.kr/Search/?stext={query}"
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code != 200:
            print(f"Failed to fetch search page, status: {response.status_code}")
            return None
            
        html = response.text
        
        # Regex to find postingCompanyName and memberSystemNo in the javascript/JSON
        # Example pattern: \"postingCompanyName\":\"㈜에스앤아이코퍼레이션\",\"userRefType\":\"C\",\"memberSystemNo\":\"46535639\"
        # Let's write a flexible regex to extract postingCompanyName and memberSystemNo
        # They might be escaped in JS strings, so we look for both escaped and non-escaped double quotes
        
        # We can find all: postingCompanyName\":\"([^\"]+)\" ... memberSystemNo\":\"(\d+)\"
        # Or memberSystemNo first, then postingCompanyName
        # Let's search all matches of memberSystemNo and find the surrounding company name
        
        candidates = []
        
        # Let's find all chunks that have postingCompanyName and memberSystemNo
        # Let's search for "postingCompanyName" in the text
        pattern = re.compile(r'postingCompanyName\\"\s*:\s*\\"\s*([^\\"]+)\s*\\".*?memberSystemNo\\"\s*:\s*\\"\s*(\d+)\s*\\"', re.DOTALL)
        for m in pattern.finditer(html):
            comp_name = m.group(1).replace('㈜', '(주)').strip()
            corp_id = m.group(2)
            candidates.append((comp_name, corp_id))
            
        # Try another regex in case order is reversed or different format
        pattern2 = re.compile(r'memberSystemNo\\"\s*:\s*\\"\s*(\d+)\s*\\".*?postingCompanyName\\"\s*:\s*\\"\s*([^\\"]+)\s*\\"', re.DOTALL)
        for m in pattern2.finditer(html):
            comp_name = m.group(2).replace('㈜', '(주)').strip()
            corp_id = m.group(1)
            candidates.append((comp_name, corp_id))
            
        # Try unescaped double quotes
        pattern3 = re.compile(r'"postingCompanyName"\s*:\s*"\s*([^"]+)\s*".*?"memberSystemNo"\s*:\s*"\s*(\d+)\s*"')
        for m in pattern3.finditer(html):
            comp_name = m.group(1).replace('㈜', '(주)').strip()
            corp_id = m.group(2)
            candidates.append((comp_name, corp_id))
            
        # De-duplicate candidates
        candidates = list(set(candidates))
        print("Candidates found in search JSON:", candidates)
        
        if not candidates:
            # Let's fallback to searching <a> tags in the HTML for Co_Read/C/
            soup = BeautifulSoup(html, 'html.parser')
            for a in soup.find_all('a', href=True):
                href = a['href']
                text = a.get_text(strip=True)
                if '/Co_Read/C/' in href:
                    m = re.search(r'/Co_Read/C/(\d+)', href)
                    if m:
                        candidates.append((text.replace('㈜', '(주)').strip(), m.group(1)))
            candidates = list(set(candidates))
            print("Candidates found in HTML <a> tags:", candidates)
            
        # Filter candidates based on name match
        clean_search_name = clean_company_name(company_name)
        matched_id = None
        for comp_name, corp_id in candidates:
            clean_cand_name = clean_company_name(comp_name)
            if clean_search_name in clean_cand_name or clean_cand_name in clean_search_name:
                print(f"Match found! Search Name: '{company_name}' -> Candidate: '{comp_name}' (ID: {corp_id})")
                matched_id = corp_id
                break
                
        if not matched_id and candidates:
            # If no perfect name match, but we have candidates, let's take the first one or the one that shares the most characters
            print("No perfect name match. Using the first candidate as fallback.")
            matched_id = candidates[0][1]
            
        return matched_id
        
    except Exception as e:
        print("Error searching:", e)
        return None

def scrape_company_details(corp_id):
    if not corp_id:
        return None
        
    url = f"https://www.jobkorea.co.kr/Recruit/Co_Read/C/{corp_id}"
    print(f"Scraping Company Details: {url}")
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        if response.status_code != 200:
            print(f"Failed to fetch company details, status: {response.status_code}")
            return None
            
        soup = BeautifulSoup(response.text, 'html.parser')
        
        result = {
            'link': url,
            '대표자': '',
            '매출액': '',
            '주소': ''
        }
        
        # 1. Parse CEO (대표자)
        # We can search inside the tables or find elements with class labels
        for th in soup.find_all(['th', 'td']):
            th_text = th.get_text(strip=True)
            if '대표자' in th_text:
                # Find sibling td or next sibling
                sibling = th.find_next_sibling()
                if sibling:
                    result['대표자'] = sibling.get_text(strip=True)
                break
                
        # 2. Parse Address (주소)
        addr_el = soup.find(class_=lambda x: x and 'address' in str(x))
        if addr_el:
            result['주소'] = addr_el.get_text(strip=True)
        else:
            # Fallback to search th/td
            for th in soup.find_all(['th', 'td']):
                th_text = th.get_text(strip=True)
                if th_text == '주소':
                    sibling = th.find_next_sibling()
                    if sibling:
                        result['주소'] = sibling.get_text(strip=True)
                    break
                    
        # 3. Parse Revenue (매출액)
        # The main company info table might have 매출액
        for th in soup.find_all(['th', 'td']):
            th_text = th.get_text(strip=True)
            if th_text == '매출액':
                sibling = th.find_next_sibling()
                if sibling:
                    result['매출액'] = sibling.get_text(strip=True)
                break
                
        # If 매출액 is not found, check the financial section or other labels
        if not result['매출액']:
            for el in soup.find_all(class_=lambda x: x and 'label' in str(x)):
                el_text = el.get_text(strip=True)
                if '매출액' in el_text:
                    parent = el.parent
                    sibling = el.find_next_sibling()
                    val = sibling.get_text(strip=True) if sibling else parent.get_text(strip=True)
                    # Extract the first matching value or clean it up
                    result['매출액'] = val
                    break
                    
        print("Scrape Result:", result)
        return result
        
    except Exception as e:
        print("Error scraping details:", e)
        return None

# Test on 맥서브
corp_id = search_jobkorea("(주)맥서브")
if corp_id:
    scrape_company_details(corp_id)
else:
    print("Could not find corp_id for 맥서브")
