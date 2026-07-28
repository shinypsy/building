# -*- coding: utf-8 -*-
"""guro_permits 폴더 월별 엑셀의 모든 시트(건축허가·착공·사용승인·임시승인) 통합."""

from __future__ import annotations

import os
import re
import sys
import warnings

import pandas as pd

warnings.filterwarnings("ignore", category=UserWarning)

sys.stdout.reconfigure(encoding="utf-8")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TARGET_DIR = os.path.join(BASE_DIR, "guro_permits")
OUTPUT_FILE = os.path.join(TARGET_DIR, "구로구_건축허가착공사용승인_통합.xlsx")
OUTPUT_USE_ONLY = os.path.join(TARGET_DIR, "구로구_사용승인현황_통합.xlsx")

SKIP_FILES = {
    "구로구 사용승인 현황(2020년이후).xlsx",
    "구로구_사용승인현황_통합.xlsx",
    "구로구_건축허가착공사용승인_통합.xlsx",
}

NTT_PATTERN = re.compile(r"^(\d+)_(\d+)_")
CATEGORY_ORDER = ["건축허가", "착공신고", "사용승인", "사용임시승인", "기타"]


def list_source_files() -> list[str]:
    files: list[str] = []
    for name in os.listdir(TARGET_DIR):
        lower = name.lower()
        if name in SKIP_FILES:
            continue
        if not lower.endswith((".xls", ".xlsx", ".xlsm")):
            continue
        if not NTT_PATTERN.match(name):
            continue
        files.append(name)
    return sorted(files)


def classify_sheet(sheet_name: str) -> str:
    text = str(sheet_name).replace(" ", "")
    if "사용(임시)승인" in text or "임시승인" in text:
        return "사용임시승인"
    if "사용승인" in text:
        return "사용승인"
    if "착공" in text:
        return "착공신고"
    if "건축허가" in text or "건축신고" in text:
        return "건축허가"
    return "기타"


def should_skip_sheet(sheet_name: str) -> bool:
    name = str(sheet_name).strip()
    if not name or name.lower() == "sheet1":
        return True
    return False


def read_excel(path: str, sheet_name: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xls":
        try:
            return pd.read_excel(path, sheet_name=sheet_name, engine="xlrd")
        except Exception:
            return pd.read_excel(path, sheet_name=sheet_name)
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(col).strip().replace("\n", "") for col in df.columns]
    return df.dropna(how="all")


def concat_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame()

    all_columns: list[str] = []
    seen: set[str] = set()
    for frame in frames:
        for col in frame.columns:
            if col not in seen:
                seen.add(col)
                all_columns.append(col)

    aligned = [frame.reindex(columns=all_columns) for frame in frames]
    return pd.concat(aligned, ignore_index=True)


def merge_all(filenames: list[str]) -> dict[str, pd.DataFrame]:
    buckets: dict[str, list[pd.DataFrame]] = {key: [] for key in CATEGORY_ORDER}
    stats = {"files": 0, "sheets": 0, "rows": 0, "skipped_files": 0}

    for filename in filenames:
        filepath = os.path.join(TARGET_DIR, filename)
        file_rows = 0
        file_sheets = 0

        try:
            workbook = pd.ExcelFile(filepath)
            for sheet_name in workbook.sheet_names:
                if should_skip_sheet(sheet_name):
                    continue

                df = read_excel(filepath, sheet_name)
                df = normalize_columns(df)
                if df.empty:
                    continue

                category = classify_sheet(sheet_name)
                df.insert(0, "출처파일명", filename)
                df.insert(1, "출처시트명", sheet_name)
                df.insert(2, "데이터구분", category)
                buckets[category].append(df)
                file_rows += len(df)
                file_sheets += 1

            if file_sheets:
                stats["files"] += 1
                stats["sheets"] += file_sheets
                stats["rows"] += file_rows
                print(f"  [OK] {filename} ({file_rows}행, 시트 {file_sheets}개)")
            else:
                stats["skipped_files"] += 1
                print(f"  [SKIP] {filename} (유효 시트 없음)")
        except Exception as exc:
            stats["skipped_files"] += 1
            print(f"  [FAIL] {filename}: {exc}")

    print(
        f"\n[SUMMARY] 파일 {stats['files']}개, 시트 {stats['sheets']}개, "
        f"총 {stats['rows']}행, 스킵/실패 {stats['skipped_files']}개"
    )

    return {key: concat_frames(buckets[key]) for key in CATEGORY_ORDER}


def write_outputs(grouped: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        for category in CATEGORY_ORDER:
            df = grouped[category]
            if df.empty:
                continue
            sheet = f"{category}_통합"[:31]
            df.to_excel(writer, index=False, sheet_name=sheet)
            print(f"  [SHEET] {sheet}: {len(df)}행, {len(df.columns)}열")

    use_frames = [
        grouped["사용승인"],
        grouped["사용임시승인"],
    ]
    use_merged = concat_frames([df for df in use_frames if not df.empty])
    if not use_merged.empty:
        with pd.ExcelWriter(OUTPUT_USE_ONLY, engine="openpyxl") as writer:
            use_merged.to_excel(writer, index=False, sheet_name="사용승인_통합")
        print(
            f"  [SHEET] 사용승인_통합(사용+임시): "
            f"{len(use_merged)}행, {len(use_merged.columns)}열"
        )


def main() -> int:
    if not os.path.isdir(TARGET_DIR):
        print(f"폴더 없음: {TARGET_DIR}")
        return 1

    filenames = list_source_files()
    if not filenames:
        print("통합할 원본 파일이 없습니다.")
        return 1

    print(f"[MERGE-ALL] 대상 파일 {len(filenames)}개\n")
    grouped = merge_all(filenames)

    if all(df.empty for df in grouped.values()):
        print("취합할 데이터가 없습니다.")
        return 1

    print(f"\n[WRITE] {OUTPUT_FILE}")
    write_outputs(grouped)
    print(f"\n[DONE]")
    print(f"  전체: {OUTPUT_FILE}")
    print(f"  사용승인+임시: {OUTPUT_USE_ONLY}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
