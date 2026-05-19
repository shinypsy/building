import sys
import pandas as pd
import re

sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel('시설관리 업체 순위.xlsx', header=None)

companies = []
for idx, row in df.iterrows():
    rank_val = str(row[0]).strip()
    name_val = str(row[1]).strip()
    
    # check if rank_val matches something like '1위', '2위', etc.
    if re.match(r'\d+위', rank_val):
        # Clean company name
        clean_name = name_val.replace('\xa0', ' ').strip()
        # Remove trailing/leading special characters if any
        companies.append({
            'rank_str': rank_val,
            'rank': int(re.findall(r'\d+', rank_val)[0]),
            'name': clean_name
        })

print(f"Extracted {len(companies)} companies.")
print("First 15 companies:")
for c in companies[:15]:
    print(c)
