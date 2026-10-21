#!/usr/bin/env python3
"""
IOPscience (JINST) Proceedings PDF Scraper

1. Reads the article list from an IOPscience proceedings page, e.g.
   https://iopscience.iop.org/journal/1748-0221/page/extraproc99
2. Looks up each article's DOI on INSPIRE-HEP and downloads the free full text
   (INSPIRE's copy, or the arXiv version).
3. Writes a CSV with a PDF_Local_Paths column.

PDFs are NOT bulk-downloaded from IOPscience itself: the site blocks automated
access and its terms forbid systematic downloading (it can get your institution's
access suspended). Papers with no free copy are listed at the end so you can grab
them by hand through your library.

Usage:
    python iop_proceedings_scraper.py
    python iop_proceedings_scraper.py https://iopscience.iop.org/journal/1748-0221/page/extraproc99

    # If the browser route gets stuck on the human-verification page:
    # open the page in your own browser, Ctrl+S "Webpage, HTML only", then:
    python iop_proceedings_scraper.py --html extraproc99.html

Requirements:
    pip install requests selenium      (selenium not needed with --html)
"""

import argparse
import csv
import html
import os
import re
import sys
import time

import requests

DEFAULT_URL = "https://iopscience.iop.org/journal/1748-0221/page/extraproc99"
INSPIRE_API = "https://inspirehep.net/api/literature"
HEADERS = {"User-Agent": "iop-proceedings-scraper/1.0 (personal research use)"}

INSPIRE_DELAY = 0.5
ARXIV_DELAY = 3.0

# JINST DOIs look like 10.1088/1748-0221/19/01/C01052
DOI_RE = re.compile(r"10\.1088/1748-0221/\d+/\d+/[A-Za-z]?\d+", re.I)
ANCHOR_RE = re.compile(r"<a\b[^>]*href=[\"']([^\"']+)[\"'][^>]*>(.*?)</a>", re.I | re.S)

CSV_HEADERS = [
    "Title", "Authors", "Journal", "Volume", "Issue", "Article", "Year",
    "DOI", "arXiv", "INSPIRE_ID", "PDF_Local_Paths",
]


# ----------------------------------------------------------------------------
# Step 1: get the article list from the IOP page
# ----------------------------------------------------------------------------

def clean_text(fragment):
    text = re.sub(r"<[^>]+>", " ", fragment)
    return re.sub(r"\s+", " ", html.unescape(text)).strip()


def parse_articles(page_html):
    """Return [(doi, title)] for every JINST article linked on the page, in page order."""
    articles = {}
    for href, inner in ANCHOR_RE.findall(page_html):
        match = DOI_RE.search(href)
        if not match:
            continue
        doi = match.group(0)
        title = clean_text(inner)
        # The same article is often linked several times (title, "PDF", "Abstract").
        # Keep the longest link text as the title.
        if doi not in articles or len(title) > len(articles[doi]):
            articles[doi] = title
    if not articles:  # fall back to bare DOIs anywhere in the page
        for doi in DOI_RE.findall(page_html):
            articles.setdefault(doi, "")
    junk = {"pdf", "abstract", "view article", "article pdf", "html", "references"}
    return [(doi, "" if t.lower() in junk else t) for doi, t in articles.items()]


def fetch_page_with_browser(url, headless):
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options

    options = Options()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1400,1000")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])

    print(f"Opening {url} in Chrome...")
    driver = webdriver.Chrome(options=options)  # Selenium 4.6+ finds chromedriver itself
    try:
        driver.get(url)
        print("Waiting for the article list (complete the 'verify you are human' "
              "check in the browser window if one appears)...")
        deadline = time.time() + 180
        while time.time() < deadline:
            page = driver.page_source
            if DOI_RE.search(page):
                time.sleep(2)  # let the rest of the list finish rendering
                return driver.page_source
            time.sleep(2)
        sys.exit("Timed out waiting for article links. Try the --html option instead "
                 "(see the top of this script).")
    finally:
        driver.quit()


# ----------------------------------------------------------------------------
# Step 2: find free full text on INSPIRE / arXiv and download it
# ----------------------------------------------------------------------------

def get_json(session, url, params=None):
    for attempt in range(4):
        resp = session.get(url, params=params, timeout=60)
        if resp.status_code == 429:
            wait = 5 * (attempt + 1)
            print(f"  Rate limited, waiting {wait}s...")
            time.sleep(wait)
            continue
        resp.raise_for_status()
        return resp.json()
    resp.raise_for_status()


def inspire_lookup(session, doi):
    params = {
        "q": f'doi:"{doi}"',
        "size": 1,
        "fields": "titles,authors.full_name,documents,arxiv_eprints,publication_info,control_number",
    }
    try:
        hits = get_json(session, INSPIRE_API, params).get("hits", {}).get("hits", [])
        return hits[0]["metadata"] if hits else {}
    except Exception as e:
        print(f"    INSPIRE lookup failed: {e}")
        return {}
    finally:
        time.sleep(INSPIRE_DELAY)


def pdf_sources(meta):
    sources = []
    for doc in meta.get("documents") or []:
        if doc.get("url") and not doc.get("hidden"):
            sources.append(("inspire", doc["url"]))
    for eprint in meta.get("arxiv_eprints") or []:
        if eprint.get("value"):
            sources.append(("arxiv", f"https://arxiv.org/pdf/{eprint['value']}"))
    return sources


def sanitize_filename(name, max_len=150):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = re.sub(r"\s+", " ", name).strip(". ")
    return name[:max_len].rstrip(". ")


def download_pdf(session, url, local_path, source):
    if os.path.exists(local_path) and os.path.getsize(local_path) > 0:
        print(f"    Already exists: {os.path.basename(local_path)}")
        return True
    print(f"    Downloading ({source}): {url}")
    tmp_path = local_path + ".part"
    try:
        resp = session.get(url, timeout=120, stream=True)
        resp.raise_for_status()
        first = b""
        with open(tmp_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if not first:
                    first = chunk
                f.write(chunk)
        if not first.startswith(b"%PDF"):
            os.remove(tmp_path)
            print("    Not a PDF (restricted copy), trying next source")
            return False
        os.replace(tmp_path, local_path)
        print(f"    Saved: {os.path.basename(local_path)}")
        return True
    except Exception as e:
        print(f"    Failed: {e}")
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        return False
    finally:
        time.sleep(ARXIV_DELAY if source == "arxiv" else INSPIRE_DELAY)


# ----------------------------------------------------------------------------

def scrape(articles, pdf_dir, output):
    os.makedirs(pdf_dir, exist_ok=True)
    session = requests.Session()
    session.headers.update(HEADERS)

    rows, missing = [], []
    for i, (doi, page_title) in enumerate(articles, start=1):
        # 10.1088/1748-0221/<volume>/<issue>/<article>
        volume, issue, artid = doi.split("/")[2:5]
        meta = inspire_lookup(session, doi)
        titles = meta.get("titles") or []
        title = (titles[0].get("title", "") if titles else "") or page_title or doi
        recid = str(meta.get("control_number", ""))
        print(f"\n[{i}/{len(articles)}] {title[:80]}")

        year = ""
        for info in meta.get("publication_info") or []:
            if info.get("year"):
                year = info["year"]
                break

        stem = f"{recid} - {title}" if recid else f"{artid} - {title}"
        local_path = os.path.join(pdf_dir, sanitize_filename(stem) + ".pdf")

        saved = False
        if not meta:
            print("    Not found on INSPIRE")
        for source, url in pdf_sources(meta):
            if download_pdf(session, url, local_path, source):
                saved = True
                break
        if not saved:
            missing.append((doi, title))

        arxivs = meta.get("arxiv_eprints") or []
        rows.append([
            title,
            "; ".join(a.get("full_name", "") for a in meta.get("authors") or []),
            "JINST", volume, issue, artid, year, doi,
            arxivs[0].get("value", "") if arxivs else "",
            recid,
            local_path if saved else "",
        ])

    with open(output, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADERS)
        writer.writerows(rows)

    print(f"\n{'=' * 60}")
    print(f"Articles found:   {len(rows)}")
    print(f"PDFs downloaded:  {len(rows) - len(missing)}")
    print(f"Without a PDF:    {len(missing)}")
    print(f"CSV file:         {output}")
    print(f"PDF directory:    {pdf_dir}")
    if missing:
        print("\nNo free copy found (open these via your library / IOPscience):")
        for doi, title in missing:
            print(f"  https://doi.org/{doi}  {title[:70]}")


def main():
    parser = argparse.ArgumentParser(description="Download PDFs for a JINST proceedings page on IOPscience")
    parser.add_argument("url", nargs="?", default=DEFAULT_URL, help="IOPscience proceedings page URL")
    parser.add_argument("--html", help="Use a saved copy of the page instead of opening a browser")
    parser.add_argument("--headless", action="store_true",
                        help="Hide the browser (often blocked by the human check)")
    parser.add_argument("--pdf-dir", default=None, help="Folder for PDFs (default: <page>_pdfs)")
    parser.add_argument("--output", default=None, help="CSV file (default: <page>.csv)")
    args = parser.parse_args()

    name = re.sub(r"\W+", "_", args.url.rstrip("/").rsplit("/", 1)[-1]) or "iop"
    if args.html:
        with open(args.html, encoding="utf-8", errors="replace") as f:
            page_html = f.read()
    else:
        page_html = fetch_page_with_browser(args.url, args.headless)

    articles = parse_articles(page_html)
    if not articles:
        sys.exit("No JINST article links found on the page.")
    print(f"Found {len(articles)} articles on the page")

    scrape(articles, args.pdf_dir or f"{name}_pdfs", args.output or f"{name}.csv")


if __name__ == "__main__":
    main()
