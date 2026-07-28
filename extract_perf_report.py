import os, PyPDF2, sys

def extract_text(pdf_path):
    with open(pdf_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        txt = ''
        for p in reader.pages:
            txt += p.extract_text() or ''
            txt += '\n'
    return txt

pdf = r'd:/Dev/Project/sample/ict/정보통신설비_2026성능점검결과서_대우테크노피아.pdf'
print('--- 추출 시작 ---')
text = extract_text(pdf)
print(text[:2000])
print('--- 끝 ---')
