import sys
import time
import pandas as pd
import scraper

sys.stdout.reconfigure(encoding='utf-8')

def get_best_corp_id(company_name):
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

def main():
    target_file = '경비경호 업체 순위_정보포함.xlsx'
    print(f"Loading {target_file}...")
    df = pd.read_excel(target_file)
    
    for idx, row in df.iterrows():
        comp_name = str(row['업체명']).strip()
        print(f"[{idx+1}/{len(df)}] Updating '{comp_name}'...")
        
        cid, details = get_best_corp_id(comp_name)
        if not details and cid:
            details = scraper.scrape_company_details(cid)
            
        if details:
            amount, date = scraper.parse_revenue_to_numeric(details.get('매출액', ''))
            
            df.at[idx, '대표자'] = details.get('대표자', 'N/A')
            import math
            df.at[idx, '매출액(원)'] = amount if amount is not None else float('nan')
            df.at[idx, '매출액 기준일'] = date if date else 'N/A'
            df.at[idx, '직원수'] = details.get('사원수', 'N/A') if details.get('사원수') else 'N/A'
            df.at[idx, '홈페이지 URL'] = details.get('홈페이지', 'N/A') if details.get('홈페이지') else 'N/A'
            df.at[idx, '사업장 주소'] = details.get('주소', 'N/A')
            df.at[idx, '잡코리아 링크'] = details.get('잡코리아링크', 'N/A')
            print(f"  -> Updated: Revenue {details.get('매출액', 'N/A')} / Address: {details.get('주소', 'N/A')}")
        else:
            print(f"  -> Failed to find better info.")
            
        time.sleep(1.0)
        
    print(f"Saving updated data to {target_file}...")
    df.to_excel(target_file, index=False)
    print("Done!")

if __name__ == '__main__':
    main()
