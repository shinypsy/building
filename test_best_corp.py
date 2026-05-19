import sys
import time
import scraper

sys.stdout.reconfigure(encoding='utf-8')

def get_best_corp_id(company_name):
    print(f"Searching best corp for: {company_name}")
    candidates = scraper.search_jobkorea(company_name)
    
    clean_search = scraper.clean_company_name(company_name)
    valid_candidates = []
    
    for comp_name, corp_id in candidates:
        clean_cand = scraper.clean_company_name(comp_name)
        if clean_search == clean_cand or clean_search in clean_cand or clean_cand in clean_search:
            valid_candidates.append((comp_name, corp_id))
            
    if not valid_candidates:
        return scraper.get_corp_id(company_name), None
        
    best_id = None
    best_revenue_val = -1
    best_details = None
    
    for c_name, c_id in valid_candidates:
        details = scraper.scrape_company_details(c_id)
        if not details:
            continue
            
        rev_val, _ = scraper.parse_revenue_to_numeric(details.get('매출액', ''))
        if rev_val is None:
            rev_val = 0
            
        if rev_val > best_revenue_val:
            best_revenue_val = rev_val
            best_id = c_id
            best_details = details
            
        time.sleep(0.5)
        
    if best_id:
        return best_id, best_details
        
    return valid_candidates[0][1], None

names = ['(주)에스원', '케이티텔레캅(주)']
for n in names:
    cid, details = get_best_corp_id(n)
    print(f"Best for {n}: {cid}")
    if details:
        print(f"  Revenue: {details['매출액']}")
        print(f"  Address: {details['주소']}")
