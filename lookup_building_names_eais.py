# -*- coding: utf-8 -*-
"""세움터(EAIS) BCIAAA02R01로 지번→건물명 조회 후 중형 시트 반영."""
from __future__ import annotations

import json
import os
import re
import sys
import time
from collections import Counter

import pandas as pd
import requests

sys.stdout.reconfigure(encoding="utf-8")

BASE = os.path.dirname(os.path.abspath(__file__))
GEUM = os.path.join(BASE, "geumcheon")
NAME_TARGETS = os.path.join(GEUM, "name_targets.json")
NAME_RESULTS = os.path.join(GEUM, "name_results_eais.json")
XLSX = os.path.join(GEUM, "금천구_사용승인현황_2016-2026.xlsx")
SHEET = "중형_연면적5천~3만"

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


def to_area(val) -> float | None:
    if val is None:
        return None
    s = str(val).strip().replace(",", "").replace(" ", "")
    if not s or s.lower() in ("nan", "none"):
        return None
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    try:
        return float(m.group(0)) if m else None
    except ValueError:
        return None


def pick_name(items: list[dict], target_area: float | None) -> tuple[str, dict]:
    """표제부 목록에서 대표 건물명 선택. 연면적 근접 우선."""
    titles = [
        it
        for it in items
        if str(it.get("regstrKindNm") or "") == "표제부" or str(it.get("regstrKindCd") or "") == "3"
    ]
    pool = titles or items

    scored = []
    for it in pool:
        name = (it.get("bldNm") or it.get("dongNm") or "").strip()
        if not name or re.fullmatch(r"\d+동", name):
            continue
        area = to_area(it.get("totArea"))
        score = 0
        if target_area is not None and area is not None:
            diff = abs(area - target_area)
            # 연면적 차이 작을수록 높은 점수
            score = 1000000 - diff
            if diff / max(target_area, 1) < 0.02:
                score += 50000
        scored.append((score, name, it))

    if not scored:
        return "", {}

    scored.sort(key=lambda x: x[0], reverse=True)
    # 상위 후보 중 이름 빈도도 고려
    top = scored[:5]
    names = [n for _, n, _ in top]
    best_name = Counter(names).most_common(1)[0][0]
    best_item = next(it for _, n, it in top if n == best_name)
    return best_name, best_item


def search_eais(dong: str, mnnm: str, slno: str) -> list[dict]:
    body = {
        "addrGbCd": "0",
        "inqireGbCd": "0",
        "bldrgstCurdiGbCd": "0",
        "bldrgstSeqno": "",
        "reqSigunguCd": SIGUNGU_CD,
        "sidoClsfCd": SIDO,
        "bjdongCd": BJDONG[dong],
        "platGbCd": "0",
        "mnnm": str(int(mnnm)),  # API는 비제로패딩 본번 사용
        "slno": str(int(slno)),
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
    data = r.json()
    items = data.get("jibunAddr") or []
    if isinstance(items, dict):
        items = [items]
    return items


def update_sheet(name_map: dict[str, str]) -> None:
    df = pd.read_excel(XLSX, sheet_name=SHEET)
    new_names = []
    changed = 0
    for _, row in df.iterrows():
        addr = short_addr(row.get("대지위치"))
        confirmed = name_map.get(addr, "")
        old = str(row.get("건축물명") or "").strip()
        if old.lower() in ("nan", "none"):
            old = ""
        if confirmed:
            new_names.append(confirmed)
            if confirmed != old:
                changed += 1
        else:
            new_names.append(old)
    df["건축물명"] = new_names
    cols = ["건축물명"] + [c for c in df.columns if c != "건축물명"]
    df = df[cols]
    with pd.ExcelWriter(XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df.to_excel(w, sheet_name=SHEET, index=False)
    print(f"[SHEET] 건축물명 갱신 {changed}/{len(df)}")


def main() -> int:
    SESSION.get("https://www.eais.go.kr/moct/bci/aaa02/BCIAAA02L01", timeout=30)

    targets = json.load(open(NAME_TARGETS, encoding="utf-8"))
    results = []
    name_map: dict[str, str] = {}
    ok = empty = fail = skip = 0

    for i, t in enumerate(targets, 1):
        addr = t["대지위치"]
        parsed = parse_jibun(addr)
        if not parsed:
            print(f"[{i}/{len(targets)}] SKIP {addr}")
            results.append({**t, "확정건물명": "", "상태": "파싱실패"})
            skip += 1
            continue
        dong, bun, ji = parsed
        target_area = to_area(t.get("연면적"))
        try:
            items = search_eais(dong, bun, ji)
            bname, item = pick_name(items, target_area)
            status = "OK" if bname else "건물명공란"
            if bname:
                ok += 1
                name_map[addr] = bname
            else:
                empty += 1
            area = item.get("totArea") if item else ""
            print(
                f"[{i}/{len(targets)}] {dong} {bun}-{ji} → {bname or '(공란)'} "
                f"(n={len(items)}, area={area})"
            )
            results.append(
                {
                    **t,
                    "확정건물명": bname,
                    "상태": status,
                    "표제부건수": len(items),
                    "매칭연면적": area,
                    "주용도_대장": item.get("mainPrposNm") if item else "",
                }
            )
        except Exception as e:
            fail += 1
            print(f"[{i}/{len(targets)}] FAIL {addr}: {e}")
            results.append({**t, "확정건물명": "", "상태": f"오류:{e}"})
        time.sleep(0.35)

    with open(NAME_RESULTS, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"[SAVE] {NAME_RESULTS}")
    print(f"요약: OK={ok} 공란={empty} 실패={fail} 스킵={skip}")

    if name_map:
        update_sheet(name_map)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
