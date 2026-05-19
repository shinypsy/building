import sys
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding='utf-8')

with open("gi_read.html", "r", encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

print("Searching for links containing company name...")
found_links = []
for a in soup.find_all('a', href=True):
    text = a.get_text(strip=True)
    if "에스앤아이" in text or "에스앤아이코퍼레이션" in text:
        found_links.append((a['href'], text))

for href, text in found_links:
    print(f"Href: {href} | Text: {text}")

print("\nAll links containing /Co_Read/ or /corp/ or /company/:")
for a in soup.find_all('a', href=True):
    href = a['href']
    if 'Co_Read' in href or 'corp' in href or 'company' in href:
        print(f"Href: {href} | Text: {a.get_text(strip=True)}")
