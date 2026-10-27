"""Scrape traffic metrics from hypestat.com/info/<domain>.

Usage:
    python hypestat_scraper.py kompas.com detik.com
    python hypestat_scraper.py -f domains.txt -o hasil.xlsx
    python hypestat_scraper.py kompas.com --browser   # use Playwright if requests gets blocked

Metrics are located by their label text (e.g. "Daily Unique Visitors") and the value that
follows it, so the scraper keeps working when the page's CSS classes change.
"""
import argparse
import random
import re
import sys
import time
from html.parser import HTMLParser

import requests

BASE_URL = "https://hypestat.com/info/{}"

# Label on the page -> column name in the output
METRICS = {
    "Daily Unique Visitors": "daily_unique_visitors",
    "Monthly Visits": "monthly_visits",
    "Pages per Visit": "pages_per_visit",
    "Daily Pageviews": "daily_pageviews",
    "Avg. Visit Duration": "avg_visit_duration",
    "Bounce Rate": "bounce_rate",
    "Global Reach": "global_reach",
    "Monthly Visits (SEMrush)": "monthly_visits_semrush",
    "Monthly Unique Visitors (SEMrush)": "monthly_unique_visitors_semrush",
    "Monthly Visits (SimilarWeb)": "monthly_visits_similarweb",
    "HypeRank": "hyperank",
    "SEMrush Rank": "semrush_rank",
    "SimilarWeb Rank": "similarweb_rank",
}

# A value looks like 1,612,878 / 2.32 / 59.37% / 04:02 / n/a
VALUE_RE = re.compile(r"^(?:[\d.,]+%?|\d{1,2}:\d{2}(?::\d{2})?|n/a|-)$", re.I)

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,id;q=0.8",
}


class _TextNodes(HTMLParser):
    """Collect visible text nodes in document order."""

    def __init__(self):
        super().__init__()
        self.nodes = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "noscript"):
            self._skip += 1

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript") and self._skip:
            self._skip -= 1

    def handle_data(self, data):
        text = " ".join(data.split())
        if text and not self._skip:
            self.nodes.append(text)


def _norm(label):
    return label.rstrip(":").strip().lower()


def parse_metrics(html):
    """Return {column: value} for every metric found in the page HTML."""
    parser = _TextNodes()
    parser.feed(html)
    nodes = parser.nodes
    wanted = {_norm(k): v for k, v in METRICS.items()}
    result = {}
    for i, node in enumerate(nodes):
        col = wanted.get(_norm(node))
        if not col or col in result:
            continue
        # The value is the first value-like text within the next few nodes
        for nxt in nodes[i + 1:i + 4]:
            if VALUE_RE.match(nxt):
                result[col] = nxt
                break
    return result


def to_number(value):
    """'1,612,878' -> 1612878, '59.37%' -> 59.37, '04:02' -> 242 (seconds)."""
    if value is None or not re.search(r"\d", value):
        return None
    if ":" in value:
        secs = 0
        for part in value.split(":"):
            secs = secs * 60 + int(part)
        return secs
    num = float(value.replace(",", "").rstrip("%"))
    return int(num) if num.is_integer() and "." not in value else num


def fetch_html(domain, session=None, browser=False):
    url = BASE_URL.format(domain)
    if browser:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch()
            page = b.new_page(user_agent=HEADERS["User-Agent"])
            page.goto(url, wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(3000)  # give a Cloudflare check time to pass
            html = page.content()
            b.close()
            return html
    resp = (session or requests).get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    return resp.text


def scrape(domain, session=None, browser=False):
    domain = re.sub(r"^https?://(www\.)?", "", domain.strip()).strip("/")
    row = {"domain": domain}
    try:
        metrics = parse_metrics(fetch_html(domain, session, browser))
        if not metrics:
            row["error"] = "no metrics found (blocked or domain not in HypeStat)"
        row.update({col: to_number(metrics.get(col)) for col in METRICS.values()})
    except Exception as e:  # keep going with the next domain
        row["error"] = f"{type(e).__name__}: {e}"
    return row


def main():
    ap = argparse.ArgumentParser(description="Scrape traffic stats from hypestat.com")
    ap.add_argument("domains", nargs="*", help="e.g. kompas.com detik.com")
    ap.add_argument("-f", "--file", help="text file with one domain per line")
    ap.add_argument("-o", "--output", default="hypestat.csv", help=".csv or .xlsx")
    ap.add_argument("--delay", type=float, default=3.0, help="seconds between requests")
    ap.add_argument("--browser", action="store_true", help="fetch with Playwright (Chromium)")
    args = ap.parse_args()

    domains = list(args.domains)
    if args.file:
        with open(args.file, encoding="utf-8") as fh:
            domains += [line.strip() for line in fh if line.strip() and not line.startswith("#")]
    if not domains:
        ap.error("give at least one domain or --file")

    import pandas as pd

    session = requests.Session()
    rows = []
    for i, domain in enumerate(domains):
        if i:
            time.sleep(args.delay + random.uniform(0, 1.5))
        row = scrape(domain, session, args.browser)
        print(f"[{i + 1}/{len(domains)}] {row['domain']}: {row.get('error') or 'ok'}", file=sys.stderr)
        rows.append(row)

    df = pd.DataFrame(rows)
    if args.output.lower().endswith(".xlsx"):
        df.to_excel(args.output, index=False)
    else:
        df.to_csv(args.output, index=False)
    print(df.to_string(index=False))
    print(f"\nSaved to {args.output}", file=sys.stderr)


if __name__ == "__main__":
    main()
