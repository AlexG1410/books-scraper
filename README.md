# Books Catalog Scraper

A Python scraper that extracts the full catalog of [Books to Scrape](https://books.toscrape.com), a sandbox site built for scraping practice, and delivers it as a clean, ready-to-use **Excel** and **CSV** file.

It walks all catalog pages automatically, cleans the data (numeric prices, 1–5 ratings, stock as numbers), removes duplicates, and formats the spreadsheet so it is ready to filter and analyze.

## What you get

| Column        | Type    | Example                   | Mode            |
|---------------|---------|---------------------------|-----------------|
| `title`       | text    | A Light in the Attic      | always          |
| `category`    | text    | Poetry                    | `--details`     |
| `price_gbp`   | number  | 51.77                     | always          |
| `rating`      | 1–5     | 3                         | always          |
| `in_stock`    | yes/no  | TRUE                      | always          |
| `stock_qty`   | number  | 22                        | `--details`     |
| `upc`         | text    | a897fe39b1053632          | `--details`     |
| `description` | text    | It's hard to imagine...   | `--details`     |
| `url`         | link    | https://books.toscrape.com/catalogue/... | always |

The Excel file comes with a styled header, frozen first row, filters on every column and price formatting.

## Features

- **Automatic pagination**: follows the "next" button until the last page, so it keeps working if the number of pages changes.
- **Two modes**: a fast listing scrape (~1 minute for 1,000 books) or a full scrape that opens every book page for extra fields.
- **Clean data**: prices as numbers, star ratings converted to 1–5, stock quantity extracted with regular expressions, duplicates removed.
- **Reliable**: automatic retries with backoff on temporary errors (429, 5xx), timeouts, and clear messages when something fails.
- **Responsible**: reads `robots.txt`, waits between requests (configurable), and identifies itself with an honest User-Agent.
- **Correct encoding**: handles the site's missing charset so symbols like `£` are never corrupted.

## Tech stack

Python 3.12+ · requests · BeautifulSoup · pandas · openpyxl

## How to run

```bash
# 1. Create and activate a virtual environment
python -m venv venv
venv\Scripts\activate          # Windows
source venv/bin/activate       # macOS / Linux

# 2. Install dependencies
pip install requests beautifulsoup4 pandas openpyxl

# 3. Run
python scraper.py                        # all books, listing data
python scraper.py --details              # all books with category, stock, UPC, description
python scraper.py --limit 20 --details   # quick test with 20 books
```

Files are saved to the `output/` folder (`books.xlsx` / `books.csv`, or `books_details.*` in details mode).

## Options

| Option        | Default  | Description                                              |
|---------------|----------|----------------------------------------------------------|
| `--details`   | off      | Open each book page for category, stock, UPC, description |
| `--limit N`   | all      | Stop after N books                                        |
| `--delay S`   | 1.0      | Seconds between requests                                  |
| `--out DIR`   | output   | Output folder                                             |

## How it works

1. Reads `robots.txt` and opens the first catalog page.
2. Extracts each book card (title, price, rating, availability, link) and follows the "next" link.
3. In details mode, visits every book page and reads the product table, breadcrumb and description.
4. Loads everything into pandas, cleans and de-duplicates it, and exports CSV and a formatted Excel file.

## Need data from another website?

I build custom scrapers that turn websites into clean spreadsheets, scheduled datasets or dashboards — including Spanish-language and Latin American sites. Only publicly available data, always respecting each site's terms.
