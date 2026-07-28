# -*- coding: utf-8 -*-
"""금천구 건축 인허가 게시판에서 사용승인현황(2016~2026) 다운로드 후 1개 파일로 취합."""
import os
import re
import sys
import time
from urllib.parse import urljoin, unquote

import pandas as pd
import requests
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding="utf-8")

BASE = "https://www.geumcheon.go.kr"
LIST_URL = f"{BASE}/portal/selectBbsNttList.do"
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "yangcheon")
RAW_DIR = os.path.join(OUT_DIR, "raw")
MERGED_PATH = os.path.join(OUT_DIR, "금천구_사용승인현황_2016-2026.xlsx")
YEAR_MIN, YEAR_MAX = 2016, 2026

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)


def extract_years(title: str) -> set[int]:
    years = set(int(y) for y in re.findall(r"(20\d{2})", title))
    # 제목에 연도가 없으면 기간 패턴에서 추출
    return years


def title_in_range(title: str) -> bool:
    years = extract_years(title)
    if not years:
        return False
    return any(YEAR_MIN <= y <= YEAR_MAX for y in years)


def is_use_approval_post(title: str) -> bool:
    t = title.replace(" ", "")
    return ("사용승인" in t) or ("사용(임시)승인" in t) or ("사용승인현황" in t)


def is_use_approval_file(name: str) -> bool:
    n = name.replace(" ", "").lower()
    keywords = ["사용승인", "사용(임시)승인", "임시승인", "사용허가"]
    return any(k in name.replace(" ", "") for k in keywords) or any(
        k in n for k in ["사용승인", "임시승인"]
    )


def parse_list_page(page_index: int, page_unit: int = 50) -> list[dict]:
    params = {
        "key": "885",
        "bbsNo": "476",
        "pageUnit": str(page_unit),
        "searchCnd": "all",
        "searchKrwd": "",
        "pageIndex": str(page_index),
    }
    r = SESSION.get(LIST_URL, params=params, timeout=60)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    posts = []
    seen = set()
    for a in soup.select('a[href*="selectBbsNttView"]'):
        href = a.get("href", "")
        m = re.search(r"nttNo=(\d+)", href)
        if not m:
            continue
        ntt_no = m.group(1)
        if ntt_no in seen:
            continue
        seen.add(ntt_no)
        title = a.get_text(" ", strip=True)
        posts.append(
            {
                "nttNo": ntt_no,
                "title": title,
                "url": urljoin(BASE, href.replace("./", "/portal/")),
            }
        )
    return posts


def collect_posts() -> list[dict]:
    all_posts = []
    page = 1
    while True:
        print(f"[LIST] page {page}")
        posts = parse_list_page(page)
        if not posts:
            break
        all_posts.extend(posts)
        # 가장 오래된 제목이 2016 이전이면 중단 가능하지만, 안전하게 끝까지
        page += 1
        if page > 80:
            break
        time.sleep(0.3)

    filtered = []
    for p in all_posts:
        if not is_use_approval_post(p["title"]):
            continue
        if not title_in_range(p["title"]):
            continue
        filtered.append(p)

    # nttNo 기준 중복 제거
    uniq = {}
    for p in filtered:
        uniq[p["nttNo"]] = p
    result = list(uniq.values())
    result.sort(key=lambda x: x["title"])
    print(f"[LIST] total matched posts: {len(result)}")
    return result


def parse_attachments(ntt_no: str) -> list[dict]:
    url = f"{BASE}/portal/selectBbsNttView.do"
    params = {"key": "885", "bbsNo": "476", "nttNo": ntt_no}
    r = SESSION.get(url, params=params, timeout=60)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")
    files = []
    for a in soup.select('a[href*="downloadBbsFile"]'):
        href = a.get("href", "")
        name = a.get_text(" ", strip=True)
        # "엑셀 문서 " 접두 제거
        name = re.sub(r"^(엑셀\s*문서|한글\s*문서|PDF\s*문서)\s*", "", name).strip()
        files.append({"name": name, "url": urljoin(BASE, href)})
    return files


def safe_filename(name: str) -> str:
    name = name.replace("\n", " ").strip()
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    return name[:180] if name else "file.xls"


def download_file(url: str, dest: str) -> bool:
    if os.path.exists(dest) and os.path.getsize(dest) > 0:
        print(f"  [SKIP] {os.path.basename(dest)}")
        return True
    r = SESSION.get(url, timeout=120)
    if r.status_code != 200 or not r.content:
        print(f"  [FAIL] {url} status={r.status_code}")
        return False
    # Content-Disposition 파일명 우선
    cd = r.headers.get("Content-Disposition", "")
    m = re.search(r"filename\*?=(?:UTF-8''|\"?)([^\";]+)", cd, re.I)
    if m and not os.path.splitext(dest)[1]:
        fname = unquote(m.group(1).strip().strip('"'))
        dest = os.path.join(os.path.dirname(dest), safe_filename(fname))
    with open(dest, "wb") as f:
        f.write(r.content)
    print(f"  [OK] {os.path.basename(dest)} ({len(r.content)} bytes)")
    return True


def read_excel_any(path: str) -> pd.DataFrame | None:
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext in (".xlsx", ".xlsm"):
            return pd.read_excel(path, engine="openpyxl")
        # .xls
        try:
            return pd.read_excel(path, engine="xlrd")
        except Exception:
            return pd.read_excel(path)
    except Exception as e:
        print(f"  [READ FAIL] {path}: {e}")
        return None


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip().replace("\n", "") for c in df.columns]
    return df


def merge_excels(file_paths: list[str]) -> pd.DataFrame:
    frames = []
    for path in file_paths:
        df = read_excel_any(path)
        if df is None or df.empty:
            continue
        df = normalize_columns(df)
        # 완전 빈 행 제거
        df = df.dropna(how="all")
        df.insert(0, "출처파일", os.path.basename(path))
        frames.append(df)
        print(f"  [MERGE] {os.path.basename(path)} rows={len(df)} cols={len(df.columns)}")

    if not frames:
        return pd.DataFrame()

    # 컬럼 합집합으로 정렬 후 concat
    all_cols = []
    seen = set()
    for f in frames:
        for c in f.columns:
            if c not in seen:
                seen.add(c)
                all_cols.append(c)
    aligned = [f.reindex(columns=all_cols) for f in frames]
    merged = pd.concat(aligned, ignore_index=True)
    return merged


def main() -> int:
    os.makedirs(RAW_DIR, exist_ok=True)
    posts = collect_posts()
    if not posts:
        print("게시글을 찾지 못했습니다.")
        return 1

    downloaded = []
    for i, post in enumerate(posts, 1):
        print(f"[{i}/{len(posts)}] {post['title']}")
        try:
            files = parse_attachments(post["nttNo"])
        except Exception as e:
            print(f"  [DETAIL FAIL] {e}")
            continue
        use_files = [f for f in files if is_use_approval_file(f["name"])]
        if not use_files:
            # 파일명에 키워드가 없으면 xls/xlsx 전부 후보로 남김 (로그)
            print(f"  [WARN] 사용승인 첨부 없음: {[f['name'] for f in files]}")
            continue
        for f in use_files:
            year_hint = "_".join(str(y) for y in sorted(extract_years(post["title"]))) or "unknown"
            base_name = safe_filename(f["name"])
            if not re.search(r"\.(xls|xlsx|xlsm)$", base_name, re.I):
                base_name += ".xls"
            dest_name = f"{post['nttNo']}_{year_hint}_{base_name}"
            dest = os.path.join(RAW_DIR, dest_name)
            if download_file(f["url"], dest):
                downloaded.append(dest)
            time.sleep(0.2)
        time.sleep(0.2)

    print(f"\n[DOWNLOAD] done: {len(downloaded)} files")
    print("[MERGE] consolidating...")
    merged = merge_excels(downloaded)
    if merged.empty:
        print("취합할 데이터가 없습니다.")
        return 1

    with pd.ExcelWriter(MERGED_PATH, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="사용승인_통합", index=False)

    print(f"[DONE] {MERGED_PATH}")
    print(f"  rows={len(merged)} cols={len(merged.columns)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
