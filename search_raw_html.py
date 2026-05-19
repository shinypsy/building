import re

with open("co_read.html", "r", encoding="utf-8") as f:
    html = f.read()

keywords = ["대표", "주소", "매출", "에스앤아이", "형형우", "서울"]
for kw in keywords:
    matches = list(re.finditer(kw, html))
    print(f"Keyword '{kw}': {len(matches)} matches")
    for m in matches[:5]:
        start = max(0, m.start() - 30)
        end = min(len(html), m.end() + 30)
        print(f"  Snippet: ... {html[start:end].strip()} ...")
