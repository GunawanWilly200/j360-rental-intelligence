# 🏢 J360 Rental Intelligence

Property price intelligence app for **SPEEDHOME.com** (Malaysia) rental
listings — built for the Jendela360 CEO Office vibe-coding test, with full
AI assistance (Claude).

**Live app:** *add your Streamlit Cloud link here*

## What it does

Enter a SPEEDHOME URL **or** search an area/property name (with autocomplete
suggestions) and the app collects live listing data and shows:

|Requirement                                                                                         |Where                                          |
|----------------------------------------------------------------------------------------------------|-----------------------------------------------|
|URL input or area search with autocomplete dropdown                                                 |Search tabs at the top                         |
|Price Summary per unit type — count, average, median, mode, fair price, avg sqft                    |📊 Price Summary table                          |
|Full unit listings — title, property, room type, RM/month, RM/year, sqft, furnishing, clickable link|🏠 Unit Listings table                          |
|Daily / Monthly / Yearly coverage with clear notice when unavailable                                |“Rent types covered” panel                     |
|Download as Excel/CSV named `SPEEDHOME_<Area>_<date>.xlsx`                                          |⬇️ Download buttons                             |
|Mobile-friendly responsive layout                                                                   |Streamlit responsive layout + scrollable tables|

**Bonus features:** 🤖 Rental Advisor (set a budget + priorities, get the
top 5 picks ranked against each segment’s fair price, with reasoning),
⚖️ area comparison mode (2–3 areas side by side with savings insights),
professional Excel export (styled sheets, live formulas that recalculate,
clickable hyperlinks, native chart — validated error-free), interactive
filters, charts, automatic market insights (incl. median RM/sqft), 1-hour
caching, self-growing autocomplete, and snapshot fallback for cloud deploys.

## How it works (technical approach)

SPEEDHOME is a Next.js site. Every search page embeds its full result set as
JSON inside `<script id="__NEXT_DATA__">` — including price, sqft, bedrooms,
bathrooms, furnishing type, and coordinates. The app extracts that JSON
directly instead of parsing fragile HTML, giving clean structured data.

- **Politeness:** the app checks `robots.txt` before scraping and waits a
  configurable delay (default 2 s) between page requests.
- **Fair Price:** median after removing outliers with the IQR method — a
  representative mid-market figure that ignores extreme listings.
- **Snapshot fallback:** every successful scrape is saved to `data/`. If the
  deployed server’s IP is blocked by SPEEDHOME, the app automatically shows
  the bundled snapshot (real scraped data), as permitted by the test brief.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy (Streamlit Community Cloud — free)

1. Push this folder to a public GitHub repo
1. Go to <https://share.streamlit.io> → “Create app”
1. Pick the repo, set **Main file path** = `app.py` → Deploy
1. Before deploying, run the app locally once for your demo areas and commit
   the generated `data/*.json` snapshots, so the cloud app always has real
   data even if SPEEDHOME blocks the host IP.

## Built with AI (vibe coding)

This project was built end-to-end with Claude: discovering the site
structure with a probe script, designing the parser around `__NEXT_DATA__`,
writing the app, and testing the statistics logic against mock data.