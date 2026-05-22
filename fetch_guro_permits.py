import os, requests, time
from bs4 import BeautifulSoup
import pandas as pd

BASE_URL = "https://www.guro.go.kr/www/selectBbsNttList.do"
PARAMS = {
    "bbsNo": "846",
    "pageUnit": "10",
    "searchCnd": "SJ",
    "searchKrwd": "건축허가",
    "key": "1871",
    "pageIndex": "1"
}
SAVE_DIR = "guro_permits"
os.makedirs(SAVE_DIR, exist_ok=True)

def get_page(page_idx):
    p = PARAMS.copy()
    p["pageIndex"] = str(page_idx)
    resp = requests.get(BASE_URL, params=p)
    resp.raise_for_status()
    return resp.text

def extract_notice_links(html):
    soup = BeautifulSoup(html, "html.parser")
    links = []
    # 각 게시글 제목 링크를 찾음 (대개 class="subject" 혹은 <a>에 "href" 포함)
    for a in soup.select('a[href*="selectBbsNttView.do"]'):
        href = a.get('href')
        if href:
            full = requests.compat.urljoin(BASE_URL, href)
            links.append(full)
    return links

def extract_excel_link(notice_url):
    resp = requests.get(notice_url)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    # 엑셀 다운로드 링크는 .xls 혹은 .xlsx 를 포함
    for a in soup.find_all('a', href=True):
        href = a['href']
        if href.lower().endswith('.xlsx') or href.lower().endswith('.xls'):
            return requests.compat.urljoin(notice_url, href)
    return None

def download_file(url, dest_path):
    with requests.get(url, stream=True) as r:
        r.raise_for_status()
        with open(dest_path, 'wb') as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)


def main():
    page = 1
    downloaded = 0
    while True:
        print(f"Fetching list page {page}...")
        html = get_page(page)
        notice_links = extract_notice_links(html)
        if not notice_links:
            print("No more notices found, stopping.")
            break
        for nl in notice_links:
            xl_url = extract_excel_link(nl)
            if xl_url:
                fname = os.path.basename(xl_url.split('?')[0])
                dest = os.path.join(SAVE_DIR, fname)
                if not os.path.exists(dest):
                    print(f"Downloading {fname}...")
                    download_file(xl_url, dest)
                    downloaded += 1
                else:
                    print(f"{fname} already exists, skipping.")
        page += 1
        time.sleep(1)  # be gentle
    print(f"Finished. Downloaded {downloaded} files.")

if __name__ == "__main__":
    main()
