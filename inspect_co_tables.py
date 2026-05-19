import sys
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding='utf-8')

with open("co_read.html", "r", encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

print("Printing meta description:")
meta_desc = soup.find('meta', {'name': 'description'})
if meta_desc:
    print(meta_desc.get('content'))

print("\nPrinting all text inside list items or tables:")
# Find table rows
rows = soup.find_all(['tr', 'li', 'div'], class_=True)
print(f"Found {len(rows)} elements with class name.")

# Let's search for terms like "대표", "매출", "주소" by lowercasing and stripping
for el in soup.find_all(['dt', 'dd', 'th', 'td', 'div', 'span', 'p']):
    text = el.get_text(strip=True)
    if not text:
        continue
    # Check if this text looks like a label
    if any(keyword in text for keyword in ["대표자", "매출액", "주소", "대표인", "본사", "기업구분", "업종"]):
        # Print the element, its siblings or parent text
        parent_text = el.parent.get_text(" | ", strip=True) if el.parent else ""
        print(f"Tag: {el.name} | Text: '{text}' | Parent/Context: '{parent_text[:200]}'")
