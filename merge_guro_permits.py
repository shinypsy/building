import os, pandas as pd
from datetime import datetime

DOWNLOAD_DIR = 'guro_permits'
OUTPUT_FILE = '구로구 건축허가 현황(2020년이후).xlsx'

if not os.path.isdir(DOWNLOAD_DIR):
    print(f'Error: directory {DOWNLOAD_DIR} does not exist')
    exit(1)

all_dfs = []
for fname in os.listdir(DOWNLOAD_DIR):
    if fname.lower().endswith(('.xls', '.xlsx')):
        path = os.path.join(DOWNLOAD_DIR, fname)
        try:
            df = pd.read_excel(path)
            df['출처 URL'] = f'https://www.guro.go.kr{fname}'  # placeholder URL
            all_dfs.append(df)
        except Exception as e:
            print(f'Failed to read {fname}: {e}')

if not all_dfs:
    print('No Excel files found')
    exit(0)

combined = pd.concat(all_dfs, ignore_index=True)
# Attempt to find a date column (common names)
date_cols = [c for c in combined.columns if '일' in str(c) or 'date' in str(c).lower()]
if date_cols:
    date_col = date_cols[0]
    combined[date_col] = pd.to_datetime(combined[date_col], errors='coerce')
    filtered = combined[combined[date_col] >= pd.Timestamp('2020-01-01')]
else:
    filtered = combined

with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    filtered.to_excel(writer, index=False, sheet_name='건축허가')
print(f'Saved merged file to {OUTPUT_FILE}, rows: {len(filtered)}')
