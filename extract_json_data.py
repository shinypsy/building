import re
import sys

sys.stdout.reconfigure(encoding='utf-8')

with open("search_result.html", "r", encoding="utf-8") as f:
    html = f.read()

# Let's search for "memberSystemNo" and print the surrounding characters
print("Searching for memberSystemNo:")
matches = list(re.finditer(r'memberSystemNo', html))
print(f"Found {len(matches)} occurrences.")
for i, m in enumerate(matches):
    start = max(0, m.start() - 100)
    end = min(len(html), m.end() + 150)
    print(f"Match {i}: ... {html[start:end]} ...")
