# -*- coding: utf-8 -*-
"""
세움터 BCIAAA02L01 검색 API로 중형시트 주소의 건축물대장(표제부) 조회 결과를
geumcheon/건축물대장/ 에 파일로 저장.

참고: 공식 PDF 발급/상세열람(BCIAZA01R02 등)은 실명인증 필요.
로그인 없이 가능한 표제부 검색(BCIAAA02R01) 결과를 JSON·엑셀로 저장함.
"""
from __future__ import annotations

import json
import os
import re
import sys
import time

import pandas as pd
import requests

sys.stdout.reconfigure(encoding="utf-8")

BASE = os.path.dirname(os.path.abspath(__file__))
GEUM = os.path.join(BASE, "geumcheon")
XLSX = os.path.join(GEUM, "금천구_사용승인현황_2016-2026.xlsx")
SHEET = "중형_연면적5천~3만"
OUT_DIR = os.path.join(GEUM, "건축물대장")
INDEX_XLSX = os.path.join(OUT_DIR, "_목록_표제부통합.xlsx")
INDEX_JSON = os.path.join(OUT_DIR, "_목록_인덱스.json")

SIGUNGU_CD = "11545"
SIDO = "11"
BJDONG = {"가산동": "10100", "독산동": "10200", "시흥동": "10300"}

SESSION = requests.Session()
SESSION.headers.update(
    {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json;charset=UTF-8",
        "Origin": "https://www.eais.go.kr",
        "Referer": "https://www.eais.go.kr/moct/bci/aaa02/BCIAAA02L01",
    }
)


def short_addr(v: str) -> str:
    s = str(v or "").strip()
    s = re.sub(r"\s*외\s*\d*\s*필지.*$", "", s)
    s = re.sub(r"\s*금천지구단위계획구역.*$", "", s)
    return s.strip()


def parse_jibun(addr: str) -> tuple[str, str, str] | None:
    s = short_addr(addr)
    m = re.search(r"(가산동|독산동|시흥동)\s*(\d+)(?:-(\d+))?", s)
    if not m:
        return None
    return m.group(1), m.group(2), m.group(3) or "0"


def safe_name(s: str) -> str:
    s = re.sub(r'[<>:"/\\|?*]', "_", str(s or "").strip())
    s = re.sub(r"\s+", "_", s)
    return s[:80] or "noname"


def search_title(dong: str, bun: str, ji: str) -> list[dict]:
    body = {
        "addrGbCd": "0",
        "inqireGbCd": "0",
        "bldrgstCurdiGbCd": "0",
        "bldrgstSeqno": "",
        "reqSigunguCd": SIGUNGU_CD,
        "sidoClsfCd": SIDO,
        "bjdongCd": BJDONG[dong],
        "platGbCd": "0",
        "mnnm": str(int(bun)),
        "slno": str(int(ji)),
        "splotNm": "",
        "blockNm": "",
        "lotNm": "",
        "roadNmCd": "",
        "bldMnnm": "",
        "bldSlno": "",
        "sigunguCd": SIGUNGU_CD,
        "untClsfCd": "",
    }
    r = SESSION.post("https://www.eais.go.kr/bci/BCIAAA02R01", json=body, timeout=40)
    r.raise_for_status()
    items = r.json().get("jibunAddr") or []
    if isinstance(items, dict):
        items = [items]
    return items


def unique_targets() -> list[dict]:
    df = pd.read_excel(XLSX, sheet_name=SHEET)
    seen = set()
    out = []
    for _, row in df.iterrows():
        addr = short_addr(row.get("대지위치"))
        if addr in seen:
            continue
        seen.add(addr)
        bname = str(row.get("건축물명") or "").strip()
        if bname.lower() in ("nan", "none"):
            bname = ""
        out.append(
            {
                "대지위치": addr,
                "건축물명": bname,
                "연면적": str(row.get("연면적(㎡)") or ""),
            }
        )
    return out


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)
    SESSION.get("https://www.eais.go.kr/moct/bci/aaa02/BCIAAA02L01", timeout=30)

    targets = unique_targets()
    print(f"대상 주소: {len(targets)}")

    index_rows = []
    all_items = []
    ok = empty = fail = skip = 0

    for i, t in enumerate(targets, 1):
        addr = t["대지위치"]
        parsed = parse_jibun(addr)
        if not parsed:
            print(f"[{i}/{len(targets)}] SKIP {addr}")
            skip += 1
            index_rows.append({**t, "상태": "파싱실패", "파일": "", "표제부건수": 0})
            continue

        dong, bun, ji = parsed
        try:
            items = search_title(dong, bun, ji)
            jibun = f"{dong}_{bun}" + (f"-{ji}" if ji != "0" else "")
            bld = ""
            if items:
                # prefer named title
                for it in items:
                    if (it.get("bldNm") or "").strip():
                        bld = it["bldNm"].strip()
                        break
                if not bld:
                    bld = (items[0].get("dongNm") or t["건축물명"] or "표제부").strip()

            fname = f"{jibun}_{safe_name(bld or t['건축물명'] or '무명')}_표제부.json"
            fpath = os.path.join(OUT_DIR, fname)

            payload = {
                "출처": "https://www.eais.go.kr/moct/bci/aaa02/BCIAAA02L01",
                "api": "/bci/BCIAAA02R01",
                "조회지번": f"{dong} {bun}-{ji}",
                "시트건축물명": t["건축물명"],
                "시트연면적": t["연면적"],
                "대지위치": addr,
                "표제부목록": items,
                "비고": "세움터 표제부 검색결과. 공식 PDF 발급/상세열람은 실명인증 필요.",
            }
            with open(fpath, "w", encoding="utf-8") as f:
                json.dump(payload, f, ensure_ascii=False, indent=2)

            # also save excel per address if items exist
            xlsx_name = fname.replace(".json", ".xlsx")
            xlsx_path = os.path.join(OUT_DIR, xlsx_name)
            if items:
                pd.DataFrame(items).to_excel(xlsx_path, index=False)
                ok += 1
            else:
                # empty placeholder excel
                pd.DataFrame([{"결과": "표제부 없음", "대지위치": addr}]).to_excel(
                    xlsx_path, index=False
                )
                empty += 1

            print(f"[{i}/{len(targets)}] {jibun} → {bld or '(공란)'} n={len(items)} | {fname}")
            index_rows.append(
                {
                    **t,
                    "상태": "OK" if items else "표제부없음",
                    "파일": fname,
                    "엑셀": xlsx_name,
                    "표제부건수": len(items),
                    "확정건물명": bld,
                }
            )
            for it in items:
                all_items.append(
                    {
                        "조회대지위치": addr,
                        "시트건축물명": t["건축물명"],
                        **it,
                    }
                )
        except Exception as e:
            fail += 1
            print(f"[{i}/{len(targets)}] FAIL {addr}: {e}")
            index_rows.append({**t, "상태": f"오류:{e}", "파일": "", "표제부건수": 0})
        time.sleep(0.35)

    # index files
    with open(INDEX_JSON, "w", encoding="utf-8") as f:
        json.dump(index_rows, f, ensure_ascii=False, indent=2)

    with pd.ExcelWriter(INDEX_XLSX, engine="openpyxl") as w:
        pd.DataFrame(index_rows).to_excel(w, sheet_name="주소별_다운로드목록", index=False)
        if all_items:
            pd.DataFrame(all_items).to_excel(w, sheet_name="표제부_전체", index=False)

    # readme
    readme = os.path.join(OUT_DIR, "README.txt")
    with open(readme, "w", encoding="utf-8") as f:
        f.write(
            "세움터(BCIAAA02L01) 건축물대장 표제부 검색결과\n"
            "출처: https://www.eais.go.kr/moct/bci/aaa02/BCIAAA02L01\n"
            "API: POST /bci/BCIAAA02R01\n\n"
            "각 주소별 *_표제부.json / *_표제부.xlsx 파일 저장.\n"
            "공식 PDF 발급·상세열람은 세움터 실명인증 후 가능합니다.\n"
        )

    print(f"\n[DONE] OUT={OUT_DIR}")
    print(f"OK={ok} 공란={empty} 실패={fail} 스킵={skip}")
    print(f"목록: {INDEX_XLSX}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
