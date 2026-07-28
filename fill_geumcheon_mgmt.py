# -*- coding: utf-8 -*-
"""네이버/구글(빙) 검색으로 관리사무소명·연락처 수집 (검증 규칙 적용)."""
from __future__ import annotations

import json
import os
import re
import sys
import time
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding="utf-8")

BASE = os.path.dirname(os.path.abspath(__file__))
GEUM = os.path.join(BASE, "geumcheon")
TARGETS = os.path.join(GEUM, "mgmt_targets_named.json")
CACHE = os.path.join(GEUM, "mgmt_lookup_cache.json")
OUT = os.path.join(GEUM, "mgmt_results_auto.json")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9",
}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)
PHONE_RE = re.compile(r"(0\d{1,2}[-.\s]?\d{3,4}[-.\s]?\d{4})")

REJECT = ("부동산", "공인중개", "중개", "분양", "매매", "임대문의", "중개사")


def normalize_phone(p: str) -> str:
    d = re.sub(r"\D", "", p)
    if d.startswith("02"):
        if len(d) == 9:
            return f"02-{d[2:5]}-{d[5:]}"
        if len(d) == 10:
            return f"02-{d[2:6]}-{d[6:]}"
    if len(d) == 10:
        return f"{d[:3]}-{d[3:6]}-{d[6:]}"
    if len(d) == 11:
        return f"{d[:3]}-{d[3:7]}-{d[7:]}"
    return p.strip()


def ok_phone(p: str) -> bool:
    d = re.sub(r"\D", "", p)
    if len(d) < 9 or len(d) > 11:
        return False
    if d.startswith("010") or d.startswith("000") or d == "0123456789":
        return False
    return d.startswith("0")


def load_cache() -> dict:
    if os.path.exists(CACHE):
        return json.load(open(CACHE, encoding="utf-8"))
    return {}


def save_cache(c: dict) -> None:
    with open(CACHE, "w", encoding="utf-8") as f:
        json.dump(c, f, ensure_ascii=False, indent=2)


def search_naver(q: str) -> list[dict]:
    r = SESSION.get(
        "https://search.naver.com/search.naver",
        params={"query": q},
        timeout=25,
    )
    html = r.text
    soup = BeautifulSoup(html, "html.parser")
    cands = []

    # place local pack
    for box in soup.select(".place_section, .api_subject_bx, .LTgSF, .E2AOU, .cXIkN"):
        text = box.get_text(" ", strip=True)
        if any(x in text for x in REJECT):
            continue
        if "관리" not in text and "전화" not in text:
            continue
        phones = [normalize_phone(p) for p in PHONE_RE.findall(text) if ok_phone(p)]
        title_el = box.select_one("a.place_bluelink, .YwYLL, .TYaxT, .OXiLu, a")
        title = title_el.get_text(" ", strip=True) if title_el else ""
        for ph in phones[:2]:
            cands.append(
                {"name": title, "phone": ph, "source": "naver", "snippet": text[:160]}
            )

    # JSON telephone in page
    for m in re.finditer(r'"telephone"\s*:\s*"([^"]+)"', html):
        ph = normalize_phone(m.group(1))
        if not ok_phone(ph):
            continue
        chunk = html[max(0, m.start() - 350) : m.end() + 20]
        nm = re.search(r'"name"\s*:\s*"([^"]{2,80})"', chunk)
        name = nm.group(1) if nm else q
        if any(x in name for x in REJECT):
            continue
        cands.append({"name": name, "phone": ph, "source": "naver_json", "snippet": ""})

    return cands


def search_bing(q: str) -> list[dict]:
    cands = []
    try:
        r = SESSION.get(
            "https://www.bing.com/search",
            params={"q": q, "setlang": "ko"},
            timeout=25,
        )
        soup = BeautifulSoup(r.text, "html.parser")
        for li in soup.select("li.b_algo"):
            text = li.get_text(" ", strip=True)
            if any(x in text for x in REJECT):
                continue
            if "관리" not in text and "전화" not in text:
                continue
            phones = [normalize_phone(p) for p in PHONE_RE.findall(text) if ok_phone(p)]
            title = ""
            h2 = li.select_one("h2")
            if h2:
                title = h2.get_text(" ", strip=True)
            for ph in phones[:1]:
                cands.append(
                    {"name": title, "phone": ph, "source": "bing", "snippet": text[:160]}
                )
    except Exception:
        pass
    return cands


def pick(bname: str, cands: list[dict]) -> dict:
    bn = re.sub(r"\s+", "", bname).lower()
    scored = []
    for c in cands:
        name = c.get("name") or ""
        snip = c.get("snippet") or ""
        blob = (name + " " + snip).replace(" ", "").lower()
        if any(x in name or x in snip for x in REJECT):
            continue
        score = 0
        if "관리사무소" in name or "관리사무소" in snip:
            score += 6
        if "관리실" in name or "관리단" in name:
            score += 4
        if bn and bn[:4] in blob:
            score += 4
        if any(x in blob for x in ("금천", "가산", "독산", "시흥", "g밸리", "디지털")):
            score += 2
        if c.get("source") == "naver_json":
            score += 2
        scored.append((score, c))
    scored.sort(key=lambda x: x[0], reverse=True)
    if not scored or scored[0][0] < 5:
        return {
            "관리사무소 업체명": "미확인",
            "관리사무소 연락처": "미확인",
            "근거": "미확인",
            "신뢰": "low",
        }
    c = scored[0][1]
    office = c.get("name") or ""
    if "관리" not in office:
        office = f"{bname} 관리사무소"
    return {
        "관리사무소 업체명": office[:80],
        "관리사무소 연락처": c["phone"],
        "근거": f"{c.get('source')}:score{scored[0][0]}",
        "신뢰": "high" if scored[0][0] >= 10 else "medium",
    }


def main() -> int:
    targets = json.load(open(TARGETS, encoding="utf-8"))
    cache = load_cache()
    results = []

    for i, t in enumerate(targets, 1):
        b = t["건축물명"]
        a = t["대지위치"]
        key = f"{b}|{a}"
        if key in cache and cache[key].get("관리사무소 연락처") not in ("", "미확인", None):
            info = cache[key]
            print(f"[{i}/{len(targets)}] CACHE {b} → {info.get('관리사무소 연락처')}")
        else:
            print(f"[{i}/{len(targets)}] {b}")
            cands = []
            for q in (f"{b} 관리사무소", f"{b} 금천구 관리사무소 전화번호"):
                try:
                    cands.extend(search_naver(q))
                except Exception as e:
                    print("  naver err", e)
                time.sleep(0.5)
                try:
                    cands.extend(search_bing(q))
                except Exception as e:
                    print("  bing err", e)
                time.sleep(0.4)
            info = pick(b, cands)
            info.update({"건축물명": b, "대지위치": a})
            cache[key] = info
            save_cache(cache)
            print(
                f"  → {info['관리사무소 업체명']} | {info['관리사무소 연락처']} | {info['근거']}"
            )
        results.append(
            {
                "건축물명": b,
                "대지위치": a,
                "관리사무소 업체명": info.get("관리사무소 업체명", "미확인"),
                "관리사무소 연락처": info.get("관리사무소 연락처", "미확인"),
                "근거": info.get("근거", ""),
                "신뢰": info.get("신뢰", ""),
            }
        )

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    filled = sum(1 for r in results if r["관리사무소 연락처"] not in ("", "미확인"))
    print(f"[DONE] filled={filled}/{len(results)} → {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
