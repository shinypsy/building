import sys
import os
import re
import time
import random
import urllib.parse
import pandas as pd
import requests
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding='utf-8')

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
}

def clean_company_name(name):
    # Remove (주), ㈜, (주 ), (유), (합), 주식회사, 등
    name = re.sub(r'[\(（]주[\)）]|㈜|[\(（]유[\)）]|[\(（]합[\)）]|주식회사', '', name)
    # Remove trailing/leading non-word characters except spaces
    name = re.sub(r'^\s+|\s+$', '', name)
    return name

def search_jobkorea(company_name):
    query = urllib.parse.quote(company_name)
    url = f"https://www.jobkorea.co.kr/Search/?stext={query}"
    
    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        if response.status_code != 200:
            return []
            
        html = response.text
        candidates = []
        
        # Method 1: Extract from JavaScript JSON in script tags
        # Search for memberSystemNo and postingCompanyName in JS JSON
        pattern1 = re.compile(r'postingCompanyName\\"\s*:\s*\\"\s*([^\\"]+)\s*\\".*?memberSystemNo\\"\s*:\s*\\"\s*(\d+)\s*\\"', re.DOTALL)
        for m in pattern1.finditer(html):
            comp_name = m.group(1).replace('㈜', '(주)').strip()
            corp_id = m.group(2)
            candidates.append((comp_name, corp_id))
            
        pattern2 = re.compile(r'memberSystemNo\\"\s*:\s*\\"\s*(\d+)\s*\\".*?postingCompanyName\\"\s*:\s*\\"\s*([^\\"]+)\s*\\"', re.DOTALL)
        for m in pattern2.finditer(html):
            comp_name = m.group(2).replace('㈜', '(주)').strip()
            corp_id = m.group(1)
            candidates.append((comp_name, corp_id))
            
        # Method 2: Extract from HTML <a> tags containing /Co_Read/C/
        soup = BeautifulSoup(html, 'html.parser')
        for a in soup.find_all('a', href=True):
            href = a['href']
            text = a.get_text(strip=True)
            if '/Co_Read/C/' in href:
                m = re.search(r'/Co_Read/C/(\d+)', href)
                if m:
                    candidates.append((text.replace('㈜', '(주)').strip(), m.group(1)))
                    
        # De-duplicate
        candidates = list(set(candidates))
        return candidates
    except Exception as e:
        print(f"  [Search Error] {e}")
        return []

def search_naver_fallback(company_name):
    query = urllib.parse.quote(f"{company_name} 잡코리아")
    url = f"https://search.naver.com/search.naver?query={query}"
    
    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        if response.status_code != 200:
            return None
            
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # Look for any links containing jobkorea.co.kr/Recruit/Co_Read
        for a in soup.find_all('a', href=True):
            href = a['href']
            if 'jobkorea.co.kr' in href and ('Co_Read' in href or 'corp' in href):
                m = re.search(r'/C/(\d+)', href)
                if m:
                    return m.group(1)
        return None
    except Exception as e:
        print(f"  [NAVER Fallback Error] {e}")
        return None

def get_corp_id(company_name):
    # Try 1: Exact search
    candidates = search_jobkorea(company_name)
    
    # Filter candidates
    clean_search = clean_company_name(company_name)
    for comp_name, corp_id in candidates:
        clean_cand = clean_company_name(comp_name)
        if clean_search == clean_cand or clean_search in clean_cand or clean_cand in clean_search:
            return corp_id
            
    # Try 2: Search with cleaned name
    if clean_search != company_name:
        candidates_clean = search_jobkorea(clean_search)
        for comp_name, corp_id in candidates_clean:
            clean_cand = clean_company_name(comp_name)
            if clean_search == clean_cand or clean_search in clean_cand or clean_cand in clean_search:
                return corp_id
                
    # Try 3: NAVER Fallback search
    naver_id = search_naver_fallback(company_name)
    if naver_id:
        return naver_id
        
    # Try 4: NAVER Fallback with cleaned name
    if clean_search != company_name:
        naver_id_clean = search_naver_fallback(clean_search)
        if naver_id_clean:
            return naver_id_clean
            
    # Try 5: If there were any candidates at all, return the first one as a last resort
    if candidates:
        return candidates[0][1]
        
    return None

def scrape_company_details(corp_id):
    if not corp_id:
        return None
        
    url = f"https://www.jobkorea.co.kr/Recruit/Co_Read/C/{corp_id}"
    
    try:
        response = requests.get(url, headers=HEADERS, timeout=10)
        if response.status_code != 200:
            return None
            
        soup = BeautifulSoup(response.text, 'html.parser')
        
        result = {
            '대표자': '',
            '매출액': '',
            '주소': '',
            '홈페이지': '',
            '사원수': '',
            '잡코리아링크': url
        }
        
        # 1. Parse CEO (대표자)
        for th in soup.find_all(['th', 'td']):
            th_text = th.get_text(strip=True)
            if '대표자' in th_text:
                sibling = th.find_next_sibling()
                if sibling:
                    result['대표자'] = sibling.get_text(strip=True)
                break
                
        # 2. Parse Address (주소)
        addr_el = soup.find(class_=lambda x: x and 'address' in str(x))
        if addr_el:
            result['주소'] = addr_el.get_text(strip=True)
        else:
            for th in soup.find_all(['th', 'td']):
                th_text = th.get_text(strip=True)
                if th_text == '주소':
                    sibling = th.find_next_sibling()
                    if sibling:
                        result['주소'] = sibling.get_text(strip=True)
                    break
                    
        # 3. Parse Revenue (매출액)
        for th in soup.find_all(['th', 'td']):
            th_text = th.get_text(strip=True)
            if th_text == '매출액':
                sibling = th.find_next_sibling()
                if sibling:
                    result['매출액'] = sibling.get_text(strip=True)
                break
                
        if not result['매출액']:
            for el in soup.find_all(class_=lambda x: x and 'label' in str(x)):
                el_text = el.get_text(strip=True)
                if '매출액' in el_text:
                    parent = el.parent
                    sibling = el.find_next_sibling()
                    val = sibling.get_text(strip=True) if sibling else parent.get_text(strip=True)
                    result['매출액'] = val
                    break
                    
        # 4. Parse Homepage (홈페이지)
        for th in soup.find_all(['th', 'td']):
            th_text = th.get_text(strip=True)
            if '홈페이지' in th_text:
                sibling = th.find_next_sibling()
                if sibling:
                    a_tag = sibling.find('a')
                    if a_tag and a_tag.get('href'):
                        result['홈페이지'] = a_tag.get('href').strip()
                    else:
                        result['홈페이지'] = sibling.get_text(strip=True)
                break
                
        # 5. Parse Employee Count (사원수)
        for th in soup.find_all(['th', 'td']):
            th_text = th.get_text(strip=True)
            if '사원수' in th_text:
                sibling = th.find_next_sibling()
                if sibling:
                    emp_text = sibling.get_text(" ", strip=True)
                    emp_clean = re.sub(r'\(.*?\)', '', emp_text).strip()
                    result['사원수'] = emp_clean
                break
                
        return result
    except Exception as e:
        print(f"  [Scrape Details Error] {e}")
        return None

def parse_korean_sub_number(s):
    s = s.strip()
    if not s:
        return 0
    if s.isdigit():
        return int(s)
        
    val = 0
    
    cheon_match = re.search(r'(\d*)천', s)
    if cheon_match:
        digits = cheon_match.group(1)
        val += (int(digits) if digits else 1) * 1000
        s = s[cheon_match.end():]
        
    baek_match = re.search(r'(\d*)백', s)
    if baek_match:
        digits = baek_match.group(1)
        val += (int(digits) if digits else 1) * 100
        s = s[baek_match.end():]
        
    sip_match = re.search(r'(\d*)십', s)
    if sip_match:
        digits = sip_match.group(1)
        val += (int(digits) if digits else 1) * 10
        s = s[sip_match.end():]
        
    if s.isdigit():
        val += int(s)
        
    return val

def parse_revenue_to_numeric(val_str):
    if not isinstance(val_str, str) or pd.isna(val_str):
        return None, None
        
    val_str = val_str.strip()
    if val_str == '-' or val_str.upper() == 'N/A' or not val_str:
        return None, None
        
    # Extract date/year in parentheses, e.g. (2025.12.31)
    date_match = re.search(r'\(([^)]+)\)', val_str)
    date_str = date_match.group(1) if date_match else None
    
    # Remove the date part
    amount_str = re.sub(r'\(([^)]+)\)', '', val_str).strip()
    # Remove commas, spaces
    amount_str = amount_str.replace(',', '').replace(' ', '')
    if amount_str.endswith('원'):
        amount_str = amount_str[:-1]
        
    total_val = 0
    
    # 1. Jo (조) part
    if '조' in amount_str:
        parts = amount_str.split('조', 1)
        jo_str = parts[0]
        total_val += parse_korean_sub_number(jo_str) * 1_000_000_000_000
        amount_str = parts[1]
        
    # 2. Eok (억) part
    if '억' in amount_str:
        parts = amount_str.split('억', 1)
        eok_str = parts[0]
        total_val += parse_korean_sub_number(eok_str) * 100_000_000
        amount_str = parts[1]
        
    # 3. Man (만) part
    if '만' in amount_str:
        parts = amount_str.split('만', 1)
        man_str = parts[0]
        total_val += parse_korean_sub_number(man_str) * 10_000
        amount_str = parts[1]
        
    # 4. Remaining (원)
    if amount_str:
        total_val += parse_korean_sub_number(amount_str)
        
    return total_val, date_str

def main():
    input_file = '시설관리 업체 순위.xlsx'
    output_file = '시설관리 업체 순위_정보포함.xlsx'
    
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    if len(sys.argv) > 2:
        output_file = sys.argv[2]
    
    if not os.path.exists(input_file):
        print(f"Error: {input_file} not found!")
        return
        
    print("Loading excel file...")
    df_raw = pd.read_excel(input_file, header=None)
    
    # Extract ranks and company names
    companies = []
    for idx, row in df_raw.iterrows():
        rank_val = str(row[0]).strip()
        name_val = str(row[1]).strip()
        
        if re.match(r'\d+위', rank_val):
            clean_name = name_val.replace('\xa0', ' ').strip()
            rank_num = int(re.findall(r'\d+', rank_val)[0])
            padded_rank = f"{rank_num:02d}위"
            companies.append({
                'rank_str': padded_rank,
                'rank': rank_num,
                'name': clean_name
            })
            
    print(f"Extracted {len(companies)} companies to search.")
    
    results = []
    total = len(companies)
    
    for i, item in enumerate(companies):
        rank_str = item['rank_str']
        name = item['name']
        
        print(f"[{i+1}/{total}] Searching '{name}' (Rank: {rank_str})...", end='', flush=True)
        
        corp_id = get_corp_id(name)
        details = None
        
        if corp_id:
            # Sleep a bit before scraping details to be polite
            time.sleep(random.uniform(0.5, 1.2))
            details = scrape_company_details(corp_id)
            
        if details:
            print(" Success!")
            amount, date = parse_revenue_to_numeric(details['매출액'])
            results.append({
                '순위': rank_str,
                '업체명': name,
                '대표자': details['대표자'],
                '매출액(원)': amount if amount is not None else '',
                '매출액 기준일': date if date else 'N/A',
                '직원수': details['사원수'] if details['사원수'] else 'N/A',
                '홈페이지 URL': details['홈페이지'] if details['홈페이지'] else 'N/A',
                '사업장 주소': details['주소'],
                '잡코리아 링크': details['잡코리아링크']
            })
        else:
            print(" Failed to find info.")
            results.append({
                '순위': rank_str,
                '업체명': name,
                '대표자': 'N/A',
                '매출액(원)': '',
                '매출액 기준일': 'N/A',
                '직원수': 'N/A',
                '홈페이지 URL': 'N/A',
                '사업장 주소': 'N/A',
                '잡코리아 링크': 'N/A'
            })
            
        # Polite delay between companies
        delay = random.uniform(1.0, 2.0)
        time.sleep(delay)
        
    # Create new DataFrame and save
    df_result = pd.DataFrame(results)
    df_result.to_excel(output_file, index=False)
    print(f"\nProcessing complete! Saved to '{output_file}'")

if __name__ == '__main__':
    main()
