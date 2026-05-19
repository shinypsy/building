import sys
import scraper

sys.stdout.reconfigure(encoding='utf-8')

candidates = scraper.search_jobkorea('(주)에스원')
print('에스원 Candidates:', candidates)

for name, cid in candidates:
    if name == '(주)에스원':
        details = scraper.scrape_company_details(cid)
        print(f'{name} ({cid}): 매출액={details["매출액"]} | 주소={details["주소"]}')
