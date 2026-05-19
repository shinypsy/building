import sys
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding='utf-8')

with open("co_read.html", "r", encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

# Company basic info is typically structured in tables with <th> and <td> or similar fields.
# Let's find all th/td pairs, or th/div pairs, or class='field-label'
print("Parsing fields:")
labels = soup.find_all(class_=lambda x: x and 'label' in x)
print(f"Found {len(labels)} label classes.")
for lbl in labels:
    # try to find the value associated with this label
    parent = lbl.parent
    sibling = lbl.find_next_sibling()
    val_text = ""
    if sibling:
        val_text = sibling.get_text(strip=True)
    else:
        # Check parent's text
        val_text = parent.get_text(" | ", strip=True) if parent else ""
    print(f"Label: '{lbl.get_text(strip=True)}' | Value context: '{val_text[:100]}'")

# Let's inspect general table rows (tr)
print("\nInspecting Table Rows:")
for tr in soup.find_all('tr'):
    th = tr.find(['th', 'td'], class_=lambda x: x and 'label' in str(x))
    td = tr.find(['td', 'div'], class_=lambda x: x and ('value' in str(x) or 'content' in str(x)))
    if not th:
        th = tr.find('th')
    if not td:
        td = tr.find('td')
    if th and td:
        print(f"Row -> Th: '{th.get_text(strip=True)}' | Td: '{td.get_text(strip=True)}'")
        
# Check for specific address class
print("\nInspecting elements with class 'address':")
for addr in soup.find_all(class_=lambda x: x and 'address' in str(x)):
    print(f"Class: '{addr.get('class')}' | Text: '{addr.get_text(strip=True)}'")
