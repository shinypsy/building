import os
import PyPDF2

def extract_text(pdf_path):
    with open(pdf_path, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        text = ''
        for page in reader.pages:
            text += page.extract_text() + '\n'
    return text

base_dir = r'd:/Dev/Project/sample/ict'
files = ['[별지 2] 정보통신설비 유지보수 관리 점검표.pdf', '[별지 3] 정보통신설비 성능점검표.pdf']
for fname in files:
    path = os.path.join(base_dir, fname)
    print(f'--- {fname} ---')
    txt = extract_text(path)
    # Print first 500 characters as preview
    print(txt[:800])
    print('... (truncated)')
