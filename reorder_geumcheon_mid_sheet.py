# -*- coding: utf-8 -*-
"""중형 시트: 건축물명 선두 배치 + 연면적 옆 관리사무소 열 추가."""
import os
import sys

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

BASE = os.path.dirname(os.path.abspath(__file__))
PATH = os.path.join(BASE, "geumcheon", "금천구_사용승인현황_2016-2026.xlsx")
SHEET = "중형_연면적5천~3만"


def main() -> int:
    df = pd.read_excel(PATH, sheet_name=SHEET)

    # 건축물명: 기존 건물명 활용
    if "건축물명" in df.columns:
        building = df["건축물명"]
    elif "건물명" in df.columns:
        building = df["건물명"]
    else:
        building = pd.Series([None] * len(df))

    df = df.copy()
    df.insert(0, "건축물명", building.values)

    # 원본 건물명 열 제거(중복)
    if "건물명" in df.columns:
        df = df.drop(columns=["건물명"])
    # insert로 생긴 중복 건축물명 방지
    arch_cols = [c for c in df.columns if c == "건축물명"]
    if len(arch_cols) > 1:
        # 첫 번째만 유지
        keep_first = True
        new_cols = []
        for c in df.columns:
            if c == "건축물명":
                if keep_first:
                    new_cols.append(c)
                    keep_first = False
                else:
                    new_cols.append("_drop_건축물명")
            else:
                new_cols.append(c)
        df.columns = new_cols
        df = df.drop(columns=[c for c in df.columns if c.startswith("_drop_")])

    area_col = next((c for c in df.columns if c == "연면적(㎡)"), None)
    if area_col is None:
        area_col = next((c for c in df.columns if "연면적" in str(c) and "증축" not in str(c)), None)
    if area_col is None:
        print("연면적 컬럼 없음")
        return 1

    # 기존 관리 열이 있으면 제거 후 연면적 옆에 재배치
    for drop_c in ("관리사무소 업체명", "관리사무소 연락처", "관리업체명", "관리업체 연락처"):
        if drop_c in df.columns:
            df = df.drop(columns=[drop_c])

    cols = list(df.columns)
    area_idx = cols.index(area_col)
    # 연면적 바로 뒤에 삽입
    front = cols[: area_idx + 1]
    back = cols[area_idx + 1 :]
    new_order = front + ["관리사무소 업체명", "관리사무소 연락처"] + back

    df["관리사무소 업체명"] = ""
    df["관리사무소 연락처"] = ""
    df = df[new_order]

    print("앞쪽 컬럼:", list(df.columns)[:12])
    print(f"행수: {len(df)}")

    with pd.ExcelWriter(PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=SHEET, index=False)

    print(f"완료: {PATH} / {SHEET}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
