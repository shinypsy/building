import sys
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding='utf-8')

with open("search_result.html", "r", encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

# Let's search for the tab "기업정보" (Corporate Info)
for el in soup.find_all(text=True):
    if "기업정보" in el:
        parent = el.parent
        print(f"Text: '{el.strip()}' | Parent tag: {parent.name} | Class: {parent.get('class')} | Href: {parent.get('href')}")
