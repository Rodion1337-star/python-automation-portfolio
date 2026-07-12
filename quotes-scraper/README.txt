============================================================
  QUOTES SCRAPER
============================================================

WHAT IT DOES
------------
The script downloads quotes from http://quotes.toscrape.com/
and saves them into an Excel file (quotes.xlsx).

What you get as the customer:
  - An Excel report with quote text, author, and tags
  - Ready-to-filter sheets (no coding needed after setup)
  - Easy control of how many pages to scrape


SETUP (once)
------------
1. Install Python 3.10+ (enable "Add to PATH" on Windows)
2. In this folder run:

   pip install requests beautifulsoup4 openpyxl


RUN
---
  python quotes_scraper.py        # first 5 pages (default)
  python quotes_scraper.py 10     # first 10 pages

Then open quotes.xlsx.


EXCEL FILE (quotes.xlsx) — 3 sheets
-----------------------------------
1) All Quotes
   One row = one quote.
   Columns: # | Quote Text | Author Name | Tag 1 | Tag 2 | ...
   (number of Tag columns depends on the data)

2) Tags Lookup
   One row = one tag linked to a quote.
   Columns: Tag Name | Author Name | Quote Text | Quote #
   Filter by "Tag Name" to find all quotes with that tag.
   "Quote #" matches "#" on the All Quotes sheet.

3) How to use
   Short tips inside the file.


TIPS
----
- Site has about 10 pages; extra pages stop automatically.
- Close quotes.xlsx before re-running (or a timestamped copy is saved).
- Delay between requests is built in (~2.5 sec).

============================================================
