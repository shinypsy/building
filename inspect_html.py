import sys
from bs4 import BeautifulSoup

sys.stdout.reconfigure(encoding='utf-8')

with open("search_result.html", "r", encoding="utf-8") as f:
    html = f.read()

soup = BeautifulSoup(html, 'html.parser')

print("Title of page:", soup.title.string if soup.title else "No title")
print("Total 'a' tags:", len(soup.find_all('a')))

# Check if there is a company list section
# Jobkorea search results page has tabs like "채용정보", "기업정보" (Corporate info)
# Let's print out some text snippets to see if we got the page correctly.
print("\nBody text snippet:")
text = soup.get_text()
lines = [line.strip() for line in text.splitlines() if line.strip()]
print("\n".join(lines[:30]))
