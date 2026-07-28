import os, re
import pandas as pd
import PyPDF2

pdf_path = r'd:/Dev/Project/sample/ict/정보통신설비_2026성능점검결과서_대우테크노피아.pdf'

# PDF 텍스트 추출
with open(pdf_path, 'rb') as f:
    reader = PyPDF2.PdfReader(f)
    full_text = ''
    for page in reader.pages:
        txt = page.extract_text()
        if txt:
            full_text += txt + '\n'

# KPI 정규식 (예시) - 실제 패턴에 맞게 수정 필요
patterns = {
    '삽입손실_dB': r'삽입손실\s*[:=]?\s*([0-9\.]+)\s*dB',
    '반사손실_dB': r'반사손실\s*[:=]?\s*([0-9\.]+)\s*dB',
    '대역폭_MHz': r'대역폭\s*[:=]?\s*([0-9\.]+)\s*MHz',
    '레이턴시_ms': r'레이턴시\s*[:=]?\s*([0-9\.]+)\s*ms',
    '패킷손실률_%': r'패킷손실률\s*[:=]?\s*([0-9\.]+)\s*%'
}

records = []
# 페이지별 혹은 라인별 매칭
for line in full_text.split('\n'):
    line = line.strip()
    if not line:
        continue
    rec = {}
    for key, pat in patterns.items():
        m = re.search(pat, line)
        if m:
            rec[key] = float(m.group(1))
    if rec:
        records.append(rec)

# DataFrame 생성 (가능한 경우)
if records:
    df = pd.DataFrame(records)
else:
    df = pd.DataFrame(columns=patterns.keys())

output_excel = r'd:/Dev/Project/sample/ict/성능점검_결과_요약.xlsx'
df.to_excel(output_excel, index=False)
print('Excel written to', output_excel)
print(df.head())
