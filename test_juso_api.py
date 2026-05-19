import requests
import json
import sys

sys.stdout.reconfigure(encoding='utf-8')

# 도로명주소 오픈 API 공용 JSONP 엔드포인트
url = 'https://www.juso.go.kr/addrlink/addrLinkApiJsonp.do'

params = {
    'confmKey': 'U01TX0FVVEhSMjAxODEwMjUxNTAzMTAxMDgyNTM=',  # 공용 데모 인증키 또는 빈 키
    'keyword': '서울특별시 양천구 신정동 1025-28',
    'resultType': 'json',
    'currentPage': '1',
    'countPerPage': '10'
}

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Referer': 'https://www.juso.go.kr/openApi/searchApi.do'
}

try:
    print("도로명주소 공용 API로 '신정동 1025-28' 검색 중...")
    res = requests.get(url, headers=headers, params=params, timeout=5)
    print(f"Status Code: {res.status_code}")
    
    text = res.text
    print(f"Raw text response length: {len(text)}")
    
    # JSONP 형식이므로 괄호 안의 JSON만 파싱
    # 형식: jsonpCallback({...})
    start_idx = text.find('(')
    end_idx = text.rfind(')')
    
    if start_idx != -1 and end_idx != -1:
        json_str = text[start_idx + 1:end_idx]
        data = json.loads(json_str)
        print("Successfully parsed JSON response!")
        
        juso_list = data.get('results', {}).get('juso', [])
        print(f"Juso count: {len(juso_list)}")
        
        for j in juso_list:
            print(f"  Road Address: {j.get('roadAddr')}")
            print(f"  Jibun Address: {j.get('jibunAddr')}")
            print(f"  Building Name (bdNm): {j.get('bdNm')}")
            
    else:
        print(f"Raw Response: {text[:500]}")
        
except Exception as e:
    print(f"Error: {e}")
