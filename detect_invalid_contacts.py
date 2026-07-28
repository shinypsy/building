import pandas as pd, re, sys, os

FILE = r'd:/Dev/Project/sample/구로구 사용승인 현황(2020년 이후).xlsx'

# Load sheet
try:
    df = pd.read_excel(FILE, sheet_name='중대형건물_관리현황')
except Exception as e:
    print('Failed to read file:', e)
    sys.exit(1)

# Define phone regex (Korean formats)
phone_pattern = re.compile(r"(\+?\d{1,2}\s*)?(\d{2,3})[-\s]?\d{3,4}[-\s]?\d{4}")

invalid_rows = []
for idx, row in df.iterrows():
    contact = str(row.get('연락처') or row.get('전화') or row.get('연락처(전화)') or '')
    if not phone_pattern.search(contact):
        invalid_rows.append((idx+2, row.get('관리사무소'), contact))  # +2 for excel row number (header+1)

if not invalid_rows:
    print('All contacts appear valid.')
else:
    print('Invalid contact rows (Excel row, 관리사무소, 현재값):')
    for r in invalid_rows:
        print(r)
    # Save a CSV for reference
    out_path = os.path.join(os.path.dirname(FILE), 'invalid_contacts.csv')
    pd.DataFrame(invalid_rows, columns=['ExcelRow','관리사무소','현재연락처']).to_csv(out_path, index=False, encoding='utf-8-sig')
    print('CSV saved to', out_path)
