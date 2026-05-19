import pandas as pd
import re
import time
import random
import sys
from scraper import get_corp_id, scrape_company_details, parse_revenue_to_numeric

sys.stdout.reconfigure(encoding='utf-8')

def parse_fm_row(row_val: str) -> tuple[str, str]:
    """'제3호 HDC랩스' 또는 ' 제4호 포스코와이드'와 같은 입력 문자열에서
    순위('제3호')와 순수 회사명('HDC랩스')을 분리하여 반환합니다.
    """
    clean_val = str(row_val).replace('\xa0', ' ').strip()
    match = re.match(r'^\s*(제\d+호)\s+(.*)$', clean_val)
    if match:
        return match.group(1), match.group(2).strip()
    return "N/A", clean_val

def run_fm_scraping() -> None:
    input_file = 'FM 업체 리스트.xlsx'
    output_file = 'FM 업체 리스트_정보포함.xlsx'
    
    print(f"[{input_file}] 파일을 읽어들이는 중...")
    df_raw = pd.read_excel(input_file, header=None)
    
    results = []
    total_rows = len(df_raw)
    print(f"총 {total_rows}개의 업체를 분석 및 수집합니다.")
    
    # 인풋 페이징 지원: 전체 인풋 리스트를 하나씩 입력받아 순차 처리
    for idx, row in df_raw.iterrows():
        raw_name = row[0]
        rep_name = row[1]
        
        rank_str, company_name = parse_fm_row(str(raw_name))
        print(f"[{idx+1}/{total_rows}] '{company_name}' ({rank_str}) 잡코리아 검색 시작...")
        
        corp_id = get_corp_id(company_name)
        details = None
        
        if corp_id:
            # 잡코리아 서버 부하 최소화를 위한 Polite delay
            time.sleep(random.uniform(0.5, 1.2))
            details = scrape_company_details(corp_id)
            
        if details:
            amount, date = parse_revenue_to_numeric(details['매출액'])
            
            # 홈페이지 주소 정제 (앞뒤 공백 제거)
            homepage = details['홈페이지'].strip() if details['홈페이지'] else 'N/A'
            if homepage.lower() == '홈페이지 없음' or not homepage:
                homepage = 'N/A'
                
            results.append({
                '순위': rank_str,
                '업체명': company_name,
                '대표자': details['대표자'] if details['대표자'] else rep_name,
                '매출액(원)': amount if amount is not None else '',
                '매출액 기준일': date if date else 'N/A',
                '직원수': details['사원수'] if details['사원수'] else 'N/A',
                '홈페이지 주소': homepage,
                '사업장주소': details['주소'] if details['주소'] else 'N/A',
                '잡코리아 링크': details['잡코리아링크']
            })
            print(f" -> 수집 성공: 대표자={details['대표자']}, 매출액={amount}, 직원수={details['사원수']}, 주소={details['주소']}")
        else:
            results.append({
                '순위': rank_str,
                '업체명': company_name,
                '대표자': rep_name,
                '매출액(원)': '',
                '매출액 기준일': 'N/A',
                '직원수': 'N/A',
                '홈페이지 주소': 'N/A',
                '사업장주소': 'N/A',
                '잡코리아 링크': 'N/A'
            })
            print(f" -> 수집 실패: 기본 정보(대표자: {rep_name})로 채웁니다.")
            
        # Polite delay
        time.sleep(random.uniform(1.0, 2.0))
        
    df_out = pd.DataFrame(results)
    df_out.to_excel(output_file, index=False)
    print(f"\n모든 작업이 완료되었습니다! 결과가 '{output_file}'에 저장되었습니다.")

if __name__ == '__main__':
    run_fm_scraping()
