import os
import pandas as pd
import warnings

# xlrd 라이브러리 경고 무시
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

TARGET_DIR = "guro_permits"
OUTPUT_FILE = "구로구 사용승인 현황(2020년이후).xlsx"

def extract_year_from_filename(filename):
    import re
    match = re.search(r'(20\d{2})', filename)
    if match:
        return int(match.group(1))
    return 0

def main():
    print(f"1. '{TARGET_DIR}' 폴더 내 2020년 이후 파일 스캔 시작...")
    
    if not os.path.exists(TARGET_DIR):
        print(f"Error: {TARGET_DIR} 폴더가 존재하지 않습니다.")
        return

    all_dfs = []
    processed_files = 0
    skipped_files = 0
    sheet_found_count = 0

    files = [f for f in os.listdir(TARGET_DIR) if f.lower().endswith(('.xls', '.xlsx'))]
    
    for filename in files:
        year = extract_year_from_filename(filename)
        
        # 2020년 이전 파일은 패스
        if year < 2020:
            skipped_files += 1
            continue
            
        filepath = os.path.join(TARGET_DIR, filename)
        
        try:
            # 엑셀 파일 로드 및 시트명 분석
            xl = pd.ExcelFile(filepath)
            target_sheets = [sheet for sheet in xl.sheet_names if '사용승인' in str(sheet)]
            
            if target_sheets:
                # 사용승인 시트가 여러 개일 경우 모두 읽어서 합침
                for sheet_name in target_sheets:
                    df = pd.read_excel(filepath, sheet_name=sheet_name)
                    # 데이터 출처 추적을 위해 소스 파일명 기입
                    df['출처파일명'] = filename 
                    df['출처시트명'] = sheet_name
                    all_dfs.append(df)
                    sheet_found_count += 1
                    
                processed_files += 1
                print(f"  -> [성공] '{filename}'에서 {len(target_sheets)}개의 사용승인 시트 추출")
            else:
                # 사용승인 시트가 없는 경우
                print(f"  -> [패스] '{filename}' (사용승인 시트 없음)")
                
        except Exception as e:
            print(f"  -> [에러] '{filename}' 읽기 실패: {e}")

    print(f"\n2. 데이터 취합 중...")
    print(f"총 스캔 대상 파일 수: {len(files)}개")
    print(f"2020년 이전 제외 파일 수: {skipped_files}개")
    print(f"사용승인 시트 추출 성공 파일 수: {processed_files}개")
    print(f"추출된 사용승인 시트 총 개수: {sheet_found_count}개")

    if not all_dfs:
        print("\n취합할 '사용승인' 데이터가 없습니다.")
        return

    # 컬럼 구조가 다를 수 있으므로 빈틈없이 병합
    final_df = pd.concat(all_dfs, ignore_index=True)
    
    print("\n3. 최종 엑셀 파일 생성 중...")
    # openpyxl 엔진으로 엑셀 저장
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        final_df.to_excel(writer, index=False, sheet_name='사용승인_통합')
        
    print(f"\n★ 대성공! 총 {len(final_df)}행의 데이터가 '{OUTPUT_FILE}' 파일에 저장되었습니다!")

if __name__ == "__main__":
    main()
