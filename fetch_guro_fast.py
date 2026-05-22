import os, requests, time
from bs4 import BeautifulSoup
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_URL = "https://www.guro.go.kr/www/selectBbsNttList.do"
PARAMS = {"bbsNo":"846","pageUnit":"10","searchCnd":"SJ","searchKrwd":"건축허가","key":"1871","pageIndex":"1"}
SAVE_DIR = "guro_permits"
os.makedirs(SAVE_DIR, exist_ok=True)

def get_page(page_idx):
    p = PARAMS.copy()
    p["pageIndex"] = str(page_idx)
    resp = requests.get(BASE_URL, params=p, timeout=10)
    resp.raise_for_status()
    return resp.text

def extract_notice_links(html):
    soup = BeautifulSoup(html, "html.parser")
    links = []
    for a in soup.select('a[href*="selectBbsNttView.do"]'):
        href = a.get('href')
        if href:
            links.append(requests.compat.urljoin(BASE_URL, href))
    return links

def extract_excel_link(notice_url):
    resp = requests.get(notice_url, timeout=10)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    for a in soup.find_all('a', href=True):
        href = a['href']
        if href.lower().endswith(('.xlsx', '.xls')):
            return requests.compat.urljoin(notice_url, href)
    return None

def download_file(url):
    fname = os.path.basename(url.split('?')[0])
    dest = os.path.join(SAVE_DIR, fname)
    if os.path.exists(dest):
        return dest
    with requests.get(url, stream=True, timeout=15) as r:
        r.raise_for_status()
        with open(dest, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
    return dest

def fetch_all():
    page = 1
    while True:
        html = get_page(page)
        notices = extract_notice_links(html)
        if not notices:
            break
        # fetch excel links concurrently
        with ThreadPoolExecutor(max_workers=8) as executor:
            future_to_url = {executor.submit(extract_excel_link, n): n for n in notices}
            for future in as_completed(future_to_url):
                xl_url = future.result()
                if xl_url:
                    download_file(xl_url)
        page += 1
        # no sleep for speed

if __name__ == "__main__":
    fetch_all()
    print('Finished fetching files.')
