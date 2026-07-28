# -*- coding: utf-8 -*-
"""자동검색 결과 정제 + 검증값 우선 반영."""
from __future__ import annotations

import json
import os
import re
import sys
from collections import Counter

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

GEUM = os.path.join(os.path.dirname(os.path.abspath(__file__)), "geumcheon")
AUTO = os.path.join(GEUM, "mgmt_results_auto.json")
VERIFIED = os.path.join(GEUM, "mgmt_verified.json")
OUT = os.path.join(GEUM, "mgmt_results_final.json")
XLSX = os.path.join(GEUM, "금천구_사용승인현황_2016-2026.xlsx")
SHEET = "중형_연면적5천~3만"

JUNK_NAME = ("k-apt", "공동주택관리정보", "새 창", "광고", "부동산", "중개", "분양")
# 서울(02) / 전국대표(15xx,16xx) / 금천 인근 허용. 043(충북) 등 타지역 다수 중복은 제외
ALLOW_PREFIX = ("02", "070", "15", "16", "18")


def short_addr(v: str) -> str:
    s = str(v or "").strip()
    s = re.sub(r"\s*외\s*\d*\s*필지.*$", "", s)
    s = re.sub(r"\s*금천지구단위계획구역.*$", "", s)
    return s.strip()


def clean_name(v: str) -> str:
    s = str(v or "").strip()
    return "" if s.lower() in ("nan", "none") else s


def is_plausible(phone: str, office: str) -> bool:
    if not phone or phone == "미확인":
        return False
    d = re.sub(r"\D", "", phone)
    if len(d) < 9 or len(d) > 11:
        return False
    if not any(d.startswith(p) or phone.startswith(p) for p in ALLOW_PREFIX):
        # 02- 형태
        if not phone.startswith("02-") and not d.startswith("02"):
            return False
    if d.startswith("010"):
        return False
    blob = (office or "").lower()
    if any(j in blob for j in JUNK_NAME):
        return False
    # 숫자만인 이상한 값
    if re.fullmatch(r"\d{8,}", phone.replace("-", "")) and not phone.startswith("02"):
        if len(re.sub(r"\D", "", phone)) >= 9 and not phone.startswith(("02", "15", "16", "070")):
            return False
    return True


def main() -> int:
    auto = json.load(open(AUTO, encoding="utf-8")) if os.path.exists(AUTO) else []
    verified = json.load(open(VERIFIED, encoding="utf-8")) if os.path.exists(VERIFIED) else {}

    # 같은 번호가 여러 건물에 반복되면 광고/오염으로 간주
    phones = [
        r.get("관리사무소 연락처")
        for r in auto
        if r.get("관리사무소 연락처") not in (None, "", "미확인")
    ]
    freq = Counter(phones)
    # 서로 다른 건물에 동일 번호가 2회 이상이면 오염으로 간주
    spam = {p for p, n in freq.items() if n >= 2}
    print("spam phones (>=2 buildings):", spam)

    cleaned = []
    for r in auto:
        b = r["건축물명"]
        a = r["대지위치"]
        office = r.get("관리사무소 업체명") or "미확인"
        phone = r.get("관리사무소 연락처") or "미확인"
        conf = r.get("신뢰") or "low"
        src = r.get("근거") or ""

        # verified by building name override
        if b in verified and verified[b].get("관리사무소 연락처") not in ("", "미확인"):
            v = verified[b]
            office, phone = v["관리사무소 업체명"], v["관리사무소 연락처"]
            conf, src = v.get("신뢰", "high"), "verified:" + v.get("근거", "")
        else:
            if phone in spam or not is_plausible(phone, office):
                office, phone, conf, src = "미확인", "미확인", "low", "필터링(오염/부적합)"

        cleaned.append(
            {
                "건축물명": b,
                "대지위치": a,
                "관리사무소 업체명": office,
                "관리사무소 연락처": phone,
                "근거": src,
                "신뢰": conf,
            }
        )

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(cleaned, f, ensure_ascii=False, indent=2)

    filled = sum(1 for r in cleaned if r["관리사무소 연락처"] != "미확인")
    print(f"정제 후 연락처 확보: {filled}/{len(cleaned)}")

    # 시트 반영
    by_key = {f"{r['건축물명']}|{short_addr(r['대지위치'])}": r for r in cleaned}
    by_name = {}
    for r in cleaned:
        by_name.setdefault(r["건축물명"], r)

    df = pd.read_excel(XLSX, sheet_name=SHEET)
    offices, phones = [], []
    hit = 0
    for _, row in df.iterrows():
        b = clean_name(row.get("건축물명"))
        a = short_addr(row.get("대지위치"))
        info = by_key.get(f"{b}|{a}") or by_name.get(b) or {}
        # verified also by name if sheet name matches
        if b in verified and verified[b].get("관리사무소 연락처") not in ("", "미확인"):
            info = {
                "관리사무소 업체명": verified[b]["관리사무소 업체명"],
                "관리사무소 연락처": verified[b]["관리사무소 연락처"],
            }
        office = info.get("관리사무소 업체명") or "미확인"
        phone = info.get("관리사무소 연락처") or "미확인"
        if phone != "미확인":
            hit += 1
        offices.append(office)
        phones.append(phone)

    df["관리사무소 업체명"] = offices
    df["관리사무소 연락처"] = phones
    cols = list(df.columns)
    # ensure 건축물명 first and office cols after 연면적
    if "건축물명" in cols:
        cols = ["건축물명"] + [c for c in cols if c != "건축물명"]
    if "연면적(㎡)" in cols and "관리사무소 업체명" in cols:
        rest = [
            c
            for c in cols
            if c not in ("건축물명", "연면적(㎡)", "관리사무소 업체명", "관리사무소 연락처")
        ]
        # keep relative order: everything before 연면적, then 연면적, office cols, rest after
        before, after = [], []
        seen_area = False
        for c in cols:
            if c in ("건축물명", "연면적(㎡)", "관리사무소 업체명", "관리사무소 연락처"):
                continue
            if not seen_area:
                # track by original position relative to area
                pass
        # simpler rebuild
        front = [c for c in cols if c not in ("관리사무소 업체명", "관리사무소 연락처")]
        if "연면적(㎡)" in front:
            i = front.index("연면적(㎡)")
            front = front[: i + 1] + ["관리사무소 업체명", "관리사무소 연락처"] + front[i + 1 :]
            # remove dup if any
            seen = set()
            newf = []
            for c in front:
                if c in seen:
                    continue
                seen.add(c)
                newf.append(c)
            front = newf
        df = df[front]

    with pd.ExcelWriter(XLSX, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df.to_excel(w, sheet_name=SHEET, index=False)

    print(f"[SHEET] 연락처 채움 {hit}/{len(df)} → {XLSX}")
    print(f"[SAVE] {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
