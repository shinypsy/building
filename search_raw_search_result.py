import re

with open("search_result.html", "r", encoding="utf-8") as f:
    html = f.read()

# Let's search for "Co_Read" or the corp_id "46511201"
for kw in ["Co_Read", "46511201", "snicorp"]:
    matches = list(re.finditer(kw, html))
    print(f"Keyword '{kw}': {len(matches)} matches")
    for m in matches[:5]:
        start = max(0, m.start() - 50)
        end = min(len(html), m.end() + 50)
        print(f"  Snippet: ... {html[start:end].strip()} ...")
