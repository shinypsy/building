# -*- coding: utf-8 -*-
"""구로구 건축과 부서자료실 - 사용승인 검색 결과 첨부파일 일괄 다운로드 (중복 제외)."""

from __future__ import annotations

import os
import re
import sys
import time
from urllib.parse import urljoin, unquote

import requests
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding="utf-8")

BASE = "https://www.guro.go.kr/www"
LIST_URL = f"{BASE}/selectBbsNttList.do"
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "guro_permits")

PARAMS = {
    "bbsNo": "846",
    "pageUnit": "10",
    "searchCnd": "SJ",
    "searchKrwd": "승인",
    "key": "1871",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)


def safe_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*]', "_", name.strip())
    return name[:160] if name else "file.xls"


def parse_list_page(page_index: int) -> list[dict]:
    params = {**PARAMS, "pageIndex": str(page_index)}
    resp = SESSION.get(LIST_URL, params=params, timeout=60)
    resp.raise_for_status()
    resp.encoding = "utf-8"
    soup = BeautifulSoup(resp.text, "html.parser")

    posts: list[dict] = []
    for row in soup.select("table tbody tr"):
        title_link = row.select_one('a[href*="selectBbsNttView"]')
        file_link = row.select_one('a[href*="downloadBbsFile"]')
        if not title_link or not file_link:
            continue

        href = title_link.get("href", "")
        file_href = file_link.get("href", "")
        if "bbsNo=846" not in href and "bbsNo=846" not in file_href:
            continue

        ntt_match = re.search(r"nttNo=(\d+)", href)
        atch_match = re.search(r"atchmnflNo=(\d+)", file_href)
        if not ntt_match or not atch_match:
            continue

        posts.append(
            {
                "nttNo": ntt_match.group(1),
                "atchmnflNo": atch_match.group(1),
                "title": title_link.get_text(" ", strip=True),
                "download_url": urljoin(BASE + "/", file_href.lstrip("/")),
            }
        )
    return posts


def collect_all_posts() -> list[dict]:
    all_posts: list[dict] = []
    seen_ntt: set[str] = set()

    page = 1
    while page <= 50:
        posts = parse_list_page(page)
        if not posts:
            break
        added = 0
        for post in posts:
            if post["nttNo"] in seen_ntt:
                continue
            seen_ntt.add(post["nttNo"])
            all_posts.append(post)
            added += 1
        print(f"[LIST] page {page}: {len(posts)}건 (신규 {added}건, 누적 {len(all_posts)}건)")
        if added == 0:
            break
        if len(posts) < int(PARAMS["pageUnit"]):
            break
        page += 1
        time.sleep(0.2)

    return all_posts


def decode_filename(content_disposition: str) -> str | None:
    if not content_disposition:
        return None
    match = re.search(r"filename\*=UTF-8''([^;]+)", content_disposition, re.I)
    if match:
        return unquote(match.group(1).strip().strip('"'))
    match = re.search(r'filename="?([^";]+)"?', content_disposition, re.I)
    if not match:
        return None
    raw = match.group(1).strip()
    for encoding in ("utf-8", "cp949", "latin-1"):
        try:
            return raw.encode("latin-1").decode(encoding)
        except (UnicodeDecodeError, UnicodeEncodeError):
            continue
    return raw


def build_dest_path(post: dict, server_name: str | None = None) -> str:
    ext = ".xlsx"
    if server_name:
        _, server_ext = os.path.splitext(server_name)
        if server_ext:
            ext = server_ext
    title_part = safe_filename(post["title"])
    filename = f"{post['nttNo']}_{post['atchmnflNo']}_{title_part}{ext}"
    return os.path.join(OUT_DIR, filename)


def load_existing_keys() -> set[str]:
    keys: set[str] = set()
    if not os.path.isdir(OUT_DIR):
        return keys
    for name in os.listdir(OUT_DIR):
        match = re.match(r"^(\d+)_(\d+)_", name)
        if match:
            keys.add(match.group(2))  # atchmnflNo
    return keys


def download_post(post: dict, seen_atch: set[str]) -> str:
    if post["atchmnflNo"] in seen_atch:
        print(f"  [SKIP-DUP] atchmnflNo={post['atchmnflNo']} | {post['title']}")
        return "dup"

    for existing in os.listdir(OUT_DIR):
        if existing.startswith(f"{post['nttNo']}_{post['atchmnflNo']}_"):
            print(f"  [SKIP-EXISTS] {existing}")
            seen_atch.add(post["atchmnflNo"])
            return "exists"

    resp = SESSION.get(post["download_url"], timeout=120)
    if resp.status_code != 200 or not resp.content:
        print(f"  [FAIL] {post['title']} status={resp.status_code}")
        return "fail"

    server_name = decode_filename(resp.headers.get("Content-Disposition", ""))
    dest = build_dest_path(post, server_name)
    with open(dest, "wb") as file:
        file.write(resp.content)

    seen_atch.add(post["atchmnflNo"])
    print(f"  [OK] {os.path.basename(dest)} ({len(resp.content)} bytes)")
    return "ok"


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)
    posts = collect_all_posts()
    if not posts:
        print("다운로드 대상 게시글이 없습니다.")
        return 1

    seen_atch = load_existing_keys()
    stats = {"ok": 0, "dup": 0, "exists": 0, "fail": 0}

    print(f"\n[DOWNLOAD] 대상 {len(posts)}건, 기존 첨부 {len(seen_atch)}건 스킵 예정\n")
    for index, post in enumerate(posts, 1):
        print(f"[{index}/{len(posts)}] {post['title']}")
        result = download_post(post, seen_atch)
        stats[result] = stats.get(result, 0) + 1
        time.sleep(0.15)

    print(
        f"\n[DONE] 저장: {OUT_DIR}\n"
        f"  신규={stats.get('ok', 0)}, 중복스킵={stats.get('dup', 0)}, "
        f"기존파일={stats.get('exists', 0)}, 실패={stats.get('fail', 0)}"
    )
    return 0 if stats.get("fail", 0) == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())
