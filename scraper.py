"""
Books to Scrape - catalog scraper

Extracts every book from https://books.toscrape.com (a sandbox site built
for scraping practice) and exports clean data to CSV and Excel.

Usage:
    python scraper.py                        # all books, listing data (~1 min)
    python scraper.py --details              # also visits each book page (~20 min)
    python scraper.py --limit 20 --details   # quick test with 20 books
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin
from urllib.robotparser import RobotFileParser

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://books.toscrape.com/"
USER_AGENT = "BooksScraper/1.0 (portfolio project)"

# The site shows ratings as a CSS class: <p class="star-rating Three">
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}

COLUMN_ORDER = [
    "title", "category", "price_gbp", "rating", "in_stock",
    "stock_qty", "upc", "description", "url",
]
EMPTY_DETAILS = {"category": None, "stock_qty": None, "upc": None, "description": None}


# --------------------------------------------------------------------------
# HTTP: polite requests with retries
# --------------------------------------------------------------------------

def make_session() -> requests.Session:
    """Session with an honest User-Agent and automatic retries on temporary errors."""
    session = requests.Session()
    session.headers["User-Agent"] = USER_AGENT
    retries = Retry(
        total=3,
        backoff_factor=1,  # waits ~1s, 2s, 4s between attempts
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


def load_robots(session: requests.Session, base_url: str) -> RobotFileParser | None:
    """Read robots.txt so the scraper never requests pages the site disallows."""
    robots_url = urljoin(base_url, "robots.txt")
    try:
        resp = session.get(robots_url, timeout=15)
    except requests.RequestException as exc:
        print(f"[warn] Could not read robots.txt ({exc}). Continuing carefully.")
        return None

    parser = RobotFileParser()
    if resp.status_code == 200:
        parser.parse(resp.text.splitlines())
        return parser
    if resp.status_code in (401, 403):
        parser.disallow_all = True
        return parser
    return None  # 404 = no robots.txt, nothing is disallowed


def fetch(session: requests.Session, url: str, delay: float,
          robots: RobotFileParser | None) -> BeautifulSoup | None:
    """Download a page and return it parsed, or None if it can't be loaded."""
    if robots is not None and not robots.can_fetch(USER_AGENT, url):
        print(f"[skip] Disallowed by robots.txt: {url}")
        return None

    time.sleep(delay)  # pause between requests so we don't overload the server
    try:
        resp = session.get(url, timeout=20)
        resp.raise_for_status()
    except requests.RequestException as exc:
        print(f"[error] {url}: {exc}")
        return None

    # The site doesn't declare its charset, so requests would guess wrong and
    # turn "£" into "Â£". Decoding the raw bytes as UTF-8 avoids that.
    return BeautifulSoup(resp.content, "html.parser", from_encoding="utf-8")


# --------------------------------------------------------------------------
# Parsing
# --------------------------------------------------------------------------

def text_of(tag) -> str:
    return tag.get_text(" ", strip=True) if tag else ""


def parse_price(text: str) -> float | None:
    """'£51.77' -> 51.77"""
    match = re.search(r"\d+(?:\.\d+)?", text.replace(",", ""))
    return float(match.group()) if match else None


def parse_rating(tag) -> int | None:
    """<p class="star-rating Three"> -> 3"""
    if tag is None:
        return None
    for css_class in tag.get("class", []):
        if css_class in RATINGS:
            return RATINGS[css_class]
    return None


def parse_listing(soup: BeautifulSoup, page_url: str) -> tuple[list[dict], str | None]:
    """Extract the books on one catalog page and the URL of the next page."""
    books = []
    for card in soup.select("article.product_pod"):
        link = card.select_one("h3 a")
        if link is None:
            continue
        books.append({
            # The visible text is truncated ("A Light in the ..."),
            # the full title lives in the title attribute.
            "title": link.get("title") or text_of(link),
            "price_gbp": parse_price(text_of(card.select_one("p.price_color"))),
            "rating": parse_rating(card.select_one("p.star-rating")),
            "in_stock": "in stock" in text_of(card.select_one("p.availability")).lower(),
            "url": urljoin(page_url, link["href"]),
        })

    next_link = soup.select_one("li.next a")
    next_url = urljoin(page_url, next_link["href"]) if next_link else None
    return books, next_url


def parse_details(soup: BeautifulSoup) -> dict:
    """Extract category, stock quantity, UPC and description from a book page."""
    table = {
        row.th.get_text(strip=True): row.td.get_text(strip=True)
        for row in soup.select("table.table tr")
        if row.th and row.td
    }

    # Breadcrumb: Home > Books > <Category> > <Title>
    crumbs = soup.select("ul.breadcrumb li a")
    category = crumbs[-1].get_text(strip=True) if len(crumbs) >= 3 else None

    # "In stock (22 available)" -> 22
    stock = re.search(r"\((\d+) available\)", table.get("Availability", ""))

    description = None
    header = soup.select_one("#product_description")
    paragraph = header.find_next_sibling("p") if header else None
    if paragraph:
        description = re.sub(r"\s*\.\.\.more$", "", paragraph.get_text(strip=True))

    return {
        "category": category,
        "stock_qty": int(stock.group(1)) if stock else None,
        "upc": table.get("UPC"),
        "description": description,
    }


# --------------------------------------------------------------------------
# Scraping flow
# --------------------------------------------------------------------------

def scrape(base_url: str, delay: float, limit: int | None, details: bool) -> list[dict]:
    session = make_session()
    robots = load_robots(session, base_url)

    # 1) Walk the catalog following the "next" button until there is none.
    url = urljoin(base_url, "catalogue/page-1.html")
    books: list[dict] = []
    seen: set[str] = set()  # skip duplicates so we never request the same book twice
    page = 0
    while url and (limit is None or len(books) < limit):
        page += 1
        soup = fetch(session, url, delay, robots)
        if soup is None:
            print("[warn] Stopping: a catalog page could not be loaded.")
            break
        page_books, url = parse_listing(soup, url)
        new_books = [b for b in page_books if b["url"] not in seen]
        seen.update(b["url"] for b in new_books)
        books.extend(new_books)
        print(f"Page {page}: {len(new_books)} new books (total {len(books)})")

    if limit is not None:
        books = books[:limit]

    # 2) Optionally open each book page for the extra fields.
    if details and books:
        print(f"\nCollecting details for {len(books)} books...")
        for i, book in enumerate(books, start=1):
            soup = fetch(session, book["url"], delay, robots)
            book.update(parse_details(soup) if soup else EMPTY_DETAILS)
            if i % 25 == 0 or i == len(books):
                print(f"Details: {i}/{len(books)}")

    return books


# --------------------------------------------------------------------------
# Cleaning and export
# --------------------------------------------------------------------------

def to_dataframe(books: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(books).drop_duplicates(subset="url")
    df = df[[c for c in COLUMN_ORDER if c in df.columns]]
    df["rating"] = df["rating"].astype("Int64")
    if "stock_qty" in df.columns:
        df["stock_qty"] = df["stock_qty"].astype("Int64")
    return df.reset_index(drop=True)


def format_sheet(ws, df: pd.DataFrame) -> None:
    """Bold header, frozen first row, filters, readable widths, price format."""
    header_fill = PatternFill("solid", fgColor="1F3A5F")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, column in enumerate(df.columns, start=1):
        sample = df[column].head(200).map(lambda v: "" if pd.isna(v) else str(v))
        longest = max([len(column)] + [len(v) for v in sample])
        ws.column_dimensions[get_column_letter(idx)].width = min(longest + 2, 60)

    if "price_gbp" in df.columns:
        letter = get_column_letter(df.columns.get_loc("price_gbp") + 1)
        for cell in ws[letter][1:]:
            cell.number_format = '"£"#,##0.00'


def export(df: pd.DataFrame, out_dir: Path, name: str) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = out_dir / f"{name}.csv"
    xlsx_path = out_dir / f"{name}.xlsx"

    # utf-8-sig lets Excel open the CSV with accents and symbols intact
    df.to_csv(csv_path, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Books")
        format_sheet(writer.sheets["Books"], df)
    return csv_path, xlsx_path


def print_summary(df: pd.DataFrame, csv_path: Path, xlsx_path: Path, seconds: float) -> None:
    ratings = df["rating"].value_counts().sort_index()
    print("\nDone!")
    print(f"  Books:          {len(df)}")
    print(f"  Average price:  {df['price_gbp'].mean():.2f} GBP")
    print(f"  In stock:       {int(df['in_stock'].sum())}")
    print("  Ratings:        " + ", ".join(f"{k} stars: {v}" for k, v in ratings.items()))
    if "category" in df.columns:
        print(f"  Categories:     {df['category'].nunique()}")
    print(f"  Time:           {seconds:.0f}s")
    print(f"  CSV:            {csv_path}")
    print(f"  Excel:          {xlsx_path}")


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Scrape the Books to Scrape catalog into clean CSV and Excel files.")
    parser.add_argument("--details", action="store_true",
                        help="also open each book page for category, stock, UPC and description (slower)")
    parser.add_argument("--limit", type=int, default=None,
                        help="maximum number of books to collect (handy for quick tests)")
    parser.add_argument("--delay", type=float, default=1.0,
                        help="seconds to wait between requests (default: 1.0)")
    parser.add_argument("--out", default="output",
                        help="output folder (default: output)")
    parser.add_argument("--base-url", default=BASE_URL,
                        help="site root; only change it to test against a local copy")
    args = parser.parse_args()

    if args.limit is not None and args.limit < 1:
        parser.error("--limit must be 1 or more")
    if args.delay < 0:
        parser.error("--delay can't be negative")
    if not args.base_url.endswith("/"):
        args.base_url += "/"
    return args


def main() -> None:
    args = parse_args()
    start = time.time()
    print(f"Scraping {args.base_url} (delay {args.delay}s between requests)\n")

    books = scrape(args.base_url, args.delay, args.limit, args.details)
    if not books:
        print("[error] No books were collected. Check your internet connection and try again.")
        sys.exit(1)

    df = to_dataframe(books)
    name = "books_details" if args.details else "books"
    try:
        csv_path, xlsx_path = export(df, Path(args.out), name)
    except PermissionError:
        print("[error] Could not save the files. If they are open in Excel, close them and run again.")
        sys.exit(1)

    print_summary(df, csv_path, xlsx_path, time.time() - start)


if __name__ == "__main__":
    main()
