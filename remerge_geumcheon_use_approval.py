# -*- coding: utf-8 -*-
"""yangcheon/raw 사용승인 xls를 헤더 보정 후 1개 파일로 재취합."""
import os
import re
import sys
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding="utf-8")

BASE = "https://www.geumcheon.go.kr"
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "yangcheon")
RAW_DIR = os.path.join(OUT_DIR, "raw")
MERGED_PATH = os.path.join(OUT_DIR, "금천구_사용승인현황_2016-2026.xlsx")
YEAR_MIN, YEAR_MAX = 2016, 2026

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    )
}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

HEADER_MARKERS = ("건축구분", "대지위치", "허가번호", "사용승인일")


def safe_filename(name: str) -> str:
    name = name.replace("\n", " ").strip()
    name = re.sub(r'[<>:"/\\|?*]', "_", name)
    return name[:180] if name else "file.xls"


def download_missing_typo_file() -> None:
    """2020.5 '시용승인현황' 오탈자 첨부 보완."""
    ntt_no = "180204"
    # find post by searching list if needed; known from warn title period
    # try discover ntt from raw missing - scan board for 2020.05
    list_url = f"{BASE}/portal/selectBbsNttList.do"
    for page in range(1, 40):
        r = SESSION.get(
            list_url,
            params={
                "key": "885",
                "bbsNo": "476",
                "pageUnit": "50",
                "pageIndex": str(page),
                "searchCnd": "all",
                "searchKrwd": "2020.05",
            },
            timeout=60,
        )
        soup = BeautifulSoup(r.text, "html.parser")
        for a in soup.select('a[href*="selectBbsNttView"]'):
            title = a.get_text(" ", strip=True)
            if "2020.05" in title.replace(" ", "") or "2020. 5" in title:
                m = re.search(r"nttNo=(\d+)", a.get("href", ""))
                if m:
                    ntt_no = m.group(1)
                    print(f"[FIX] found 2020.05 post nttNo={ntt_no} title={title}")
                    _download_siyong(ntt_no, title)
                    return
        time.sleep(0.2)
    print("[FIX] 2020.05 post not found via search; trying direct scan")
    # fallback: open known warn period posts by scanning downloads missing
    for page in range(1, 40):
        r = SESSION.get(
            list_url,
            params={"key": "885", "bbsNo": "476", "pageUnit": "50", "pageIndex": str(page)},
            timeout=60,
        )
        soup = BeautifulSoup(r.text, "html.parser")
        for a in soup.select('a[href*="selectBbsNttView"]'):
            title = a.get_text(" ", strip=True)
            compact = title.replace(" ", "")
            if "2020.05.01" in compact or "2020.5.1" in compact.replace(" ", ""):
                m = re.search(r"nttNo=(\d+)", a.get("href", ""))
                if m:
                    _download_siyong(m.group(1), title)
                    return
        time.sleep(0.2)


def _download_siyong(ntt_no: str, title: str) -> None:
    url = f"{BASE}/portal/selectBbsNttView.do"
    r = SESSION.get(url, params={"key": "885", "bbsNo": "476", "nttNo": ntt_no}, timeout=60)
    soup = BeautifulSoup(r.text, "html.parser")
    for a in soup.select('a[href*="downloadBbsFile"]'):
        name = a.get_text(" ", strip=True)
        name = re.sub(r"^(엑셀\s*문서|한글\s*문서|PDF\s*문서)\s*", "", name).strip()
        if "시용승인" in name.replace(" ", "") or "사용승인" in name.replace(" ", ""):
            dest = os.path.join(
                RAW_DIR, f"{ntt_no}_2020_{safe_filename(name if name.endswith('.xls') else name + '.xls')}"
            )
            if os.path.exists(dest) and os.path.getsize(dest) > 0:
                print(f"[FIX] already exists {dest}")
                return
            fr = SESSION.get(urljoin(BASE, a["href"]), timeout=120)
            with open(dest, "wb") as f:
                f.write(fr.content)
            print(f"[FIX] downloaded {dest} ({len(fr.content)} bytes)")
            return
    print(f"[FIX] no matching attachment on nttNo={ntt_no}")


def is_cumulative_dump(filename: str) -> bool:
    """다년간 누적 조회 파일 제외 (월/주간 단위만 유지)."""
    name = filename
    # 2014.01.01~2017 / 2015.01.01~2018 같은 누적본
    if re.search(r"2014\.01\.01\s*~?\s*2017", name):
        return True
    if re.search(r"2015\.01\.01\s*~?\s*2018", name):
        return True
    if "2014_2017" in name or "2015_2018" in name:
        return True
    return False


def find_header_row(df: pd.DataFrame) -> int:
    for i in range(min(10, len(df))):
        row = [str(v).strip() for v in df.iloc[i].tolist()]
        joined = "|".join(row)
        if "건축구분" in joined and ("대지위치" in joined or "연면적" in joined):
            return i
        hit = sum(1 for m in HEADER_MARKERS if m in joined)
        if hit >= 2:
            return i
    return 0


def read_excel_smart(path: str) -> pd.DataFrame | None:
    try:
        raw = pd.read_excel(path, header=None, dtype=object)
    except Exception as e:
        print(f"  [READ FAIL] {os.path.basename(path)}: {e}")
        return None
    if raw.empty:
        return None
    h = find_header_row(raw)
    headers = [str(c).strip().replace("\n", "") for c in raw.iloc[h].tolist()]
    # Unnamed / nan 헤더 정리
    clean_headers = []
    for i, c in enumerate(headers):
        if c.lower() in ("nan", "none", "") or c.startswith("Unnamed"):
            clean_headers.append(f"컬럼{i}")
        else:
            clean_headers.append(c)
    df = raw.iloc[h + 1 :].copy()
    df.columns = clean_headers
    df = df.dropna(how="all")
    # 중복 컬럼명 먼저 정리
    cols = list(df.columns)
    seen_cols: dict[str, int] = {}
    new_cols: list[str] = []
    for c in cols:
        if c not in seen_cols:
            seen_cols[c] = 0
            new_cols.append(c)
        else:
            seen_cols[c] += 1
            new_cols.append(f"{c}_{seen_cols[c]}")
    df.columns = new_cols
    # 헤더가 데이터로 반복된 행 제거
    if "건축구분" in df.columns:
        col = df["건축구분"]
        if isinstance(col, pd.DataFrame):
            col = col.iloc[:, 0]
        df = df[col.map(lambda x: str(x).strip() != "건축구분")]
    return df.reset_index(drop=True)


def normalize_col_name(c: str) -> str:
    c = str(c).strip().replace("\n", "").replace(" ", "")
    mapping = {
        "대지면적(m2)": "대지면적(㎡)",
        "건축면적(m2)": "건축면적(㎡)",
        "연면적(m2)": "연면적(㎡)",
        "증축연면적(m2)": "증축연면적(㎡)",
        "총주차장면적(m2)": "총주차장면적(㎡)",
        "최고높이(M)": "최고높이(m)",
        "세대가구수": "세대수",
        "총주차": "주차장대수",
    }
    return mapping.get(c, c)


def parse_year_from_date(val) -> int | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()
    m = re.search(r"(20\d{2})", s)
    return int(m.group(1)) if m else None


def main() -> int:
    os.makedirs(RAW_DIR, exist_ok=True)
    print("[FIX] try download typo attachment...")
    try:
        download_missing_typo_file()
    except Exception as e:
        print(f"[FIX] skip: {e}")

    files = [
        os.path.join(RAW_DIR, n)
        for n in sorted(os.listdir(RAW_DIR))
        if n.lower().endswith((".xls", ".xlsx", ".xlsm")) and not is_cumulative_dump(n)
    ]
    print(f"[MERGE] candidate files: {len(files)}")

    frames = []
    for path in files:
        df = read_excel_smart(path)
        if df is None or df.empty:
            continue
        df = df.rename(columns={c: normalize_col_name(c) for c in df.columns})
        # 중복 컬럼명 처리
        cols = list(df.columns)
        seen = {}
        new_cols = []
        for c in cols:
            if c not in seen:
                seen[c] = 0
                new_cols.append(c)
            else:
                seen[c] += 1
                new_cols.append(f"{c}_{seen[c]}")
        df.columns = new_cols

        df.insert(0, "출처파일", os.path.basename(path))
        # 사용승인일 기준 연도 필터 (가능하면)
        if "사용승인일" in df.columns:
            years = df["사용승인일"].map(parse_year_from_date)
            mask = years.map(lambda y: y is None or (YEAR_MIN <= y <= YEAR_MAX))
            before = len(df)
            df = df[mask].copy()
            # None은 유지하되, 확실한 범위 밖만 제거됨
        frames.append(df)
        print(f"  [OK] {os.path.basename(path)} rows={len(df)}")

    if not frames:
        print("no data")
        return 1

    all_cols = []
    seen = set()
    for f in frames:
        for c in f.columns:
            if c not in seen:
                seen.add(c)
                all_cols.append(c)
    # 의미 없는 컬럼N 은 뒤로
    primary = [c for c in all_cols if not re.fullmatch(r"컬럼\d+", c)]
    misc = [c for c in all_cols if re.fullmatch(r"컬럼\d+", c)]
    ordered = primary + misc
    merged = pd.concat([f.reindex(columns=ordered) for f in frames], ignore_index=True)

    # 완전 빈 데이터 행 제거 (출처파일만 있는 경우)
    data_cols = [c for c in merged.columns if c != "출처파일"]
    merged = merged.dropna(subset=data_cols, how="all")

    with pd.ExcelWriter(MERGED_PATH, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="사용승인_통합", index=False)

    print(f"[DONE] {MERGED_PATH}")
    print(f"  rows={len(merged)} cols={len(merged.columns)} files={len(frames)}")
    if "사용승인일" in merged.columns:
        ys = merged["사용승인일"].map(parse_year_from_date).value_counts(dropna=False).sort_index()
        print("  year counts:")
        print(ys.to_string())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
