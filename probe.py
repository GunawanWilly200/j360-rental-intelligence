"""
SPEEDHOME structure probe.
Run:  python probe.py
Then upload the generated files back to Claude:
  - probe_report.txt
  - speedhome_search.html  (only if asked; it can be large)
"""
import json
import re

import cloudscraper

URL = "https://speedhome.com/rent/mont-kiara"

scraper = cloudscraper.create_scraper()
html = scraper.get(URL).text

with open("speedhome_search.html", "w", encoding="utf-8") as f:
    f.write(html)

report = []
report.append(f"URL: {URL}")
report.append(f"HTML length: {len(html):,} chars")

# 1. Which app framework / state object?
for marker in ["__NUXT__", "__NEXT_DATA__", "__INITIAL_STATE__",
               "application/ld+json", "window.dataLayer"]:
    pos = html.find(marker)
    report.append(f"Marker {marker!r}: {'FOUND at ' + str(pos) if pos != -1 else 'not found'}")

# 2. API endpoints referenced in the page
apis = sorted(set(re.findall(r'https?://[^"\'\s]*api[^"\'\s]*', html)))[:30]
report.append("\n--- API-looking URLs found in page ---")
report.extend(apis or ["(none)"])

rel_apis = sorted(set(re.findall(r'["\'](/(?:api|graphql|search)[^"\']{0,120})["\']', html)))[:30]
report.append("\n--- Relative API paths ---")
report.extend(r[0] if isinstance(r, tuple) else r for r in rel_apis) if rel_apis else report.append("(none)")

# 3. Sample context around price-like JSON keys
report.append("\n--- Context around price keys (first 5 hits each) ---")
for key in ['"price"', '"rent"', '"monthlyRent"', '"sqft"', '"furnish']:
    hits = [m.start() for m in re.finditer(re.escape(key), html)][:5]
    report.append(f"\n{key}: {len(hits)} shown")
    for h in hits:
        snippet = html[max(0, h - 80): h + 220].replace("\n", " ")
        report.append(f"  …{snippet}…")

# 4. JSON-LD blocks (often contain clean listing data)
ld_blocks = re.findall(
    r'<script[^>]*application/ld\+json[^>]*>(.*?)</script>', html, re.S)[:3]
report.append(f"\n--- JSON-LD blocks: {len(ld_blocks)} found ---")
for i, block in enumerate(ld_blocks):
    try:
        pretty = json.dumps(json.loads(block), indent=1)[:1500]
    except json.JSONDecodeError:
        pretty = block[:800]
    report.append(f"\n[JSON-LD #{i}]\n{pretty}")

# 5. Count how many listing links appear
links = sorted(set(re.findall(r'href="(/details/[^"]+)"', html)))
report.append(f"\n--- /details/ links on page: {len(links)} ---")
report.extend(links[:15])

with open("probe_report.txt", "w", encoding="utf-8") as f:
    f.write("\n".join(report))

print("\n".join(report[:25]))
print("\n✅ Done! Files created: probe_report.txt and speedhome_search.html")
print("👉 Upload probe_report.txt back to Claude.")
