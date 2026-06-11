import cloudscraper
import re
import json

url = "https://speedhome.com/details/residensi-suasana-damai-howqjhqa"

scraper = cloudscraper.create_scraper()

html = scraper.get(url).text

match = re.search(
    r'"sqft":(\d+).*?"bedroom":(\d+).*?"bathroom":(\d+).*?"carpark":(\d+).*?"price":(\d+)',
    html
)

if match:
    print("SQFT:", match.group(1))
    print("BEDROOM:", match.group(2))
    print("BATHROOM:", match.group(3))
    print("PARKING:", match.group(4))
    print("PRICE:", match.group(5))
else:
    print("Not found")