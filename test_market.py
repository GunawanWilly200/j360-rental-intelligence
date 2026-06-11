import cloudscraper
import re

url = "https://speedhome.com/rent/mont-kiara"

scraper = cloudscraper.create_scraper()

html = scraper.get(url).text

prices = re.findall(r'RM\s?\d+', html)

print(prices[:20])