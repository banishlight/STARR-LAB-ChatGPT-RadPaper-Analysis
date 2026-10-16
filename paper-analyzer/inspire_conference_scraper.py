#!/usr/bin/env python3
"""
INSPIRE-HEP Conference PDF Scraper

Collects every paper linked to an INSPIRE conference record, downloads the
PDFs, and writes a CSV with a column pointing at each local PDF.

Uses INSPIRE's public JSON API, so no browser/Selenium is needed.

Usage:
    python inspire_conference_scraper.py                      # TWEPP23 (2639288)
    python inspire_conference_scraper.py 2639288
    python inspire_conference_scraper.py https://inspirehep.net/conferences/2639288
    python inspire_conference_scraper.py 2639288 --pdf-dir twepp23_pdfs --output twepp23.csv

Requirements:
    pip install requests
"""

import argparse
import csv
import os
import re
import sys
import time

import requests

API = "https://inspirehep.net/api"
HEADERS = {"User-Agent": "inspire-conference-scraper/1.0 (personal research use)"}

# INSPIRE allows ~15 requests per 5 seconds; arXiv asks for ~3 s between requests.
INSPIRE_DELAY = 0.5
ARXIV_DELAY = 3.0

CSV_HEADERS = [
    "Title", "Authors", "Journal", "Volume", "Article", "Year",
    "DOI", "arXiv", "INSPIRE_ID", "PDF_Local_Paths",
]


def parse_conference_id(value):
    """Accept a bare ID or an inspirehep.net conference URL."""
    match = re.search(r"(\d+)\s*/?\s*$", value.strip())
    if not match:
        sys.exit(f"Could not find a conference ID in: {value}")
    return match.group(1)


def sanitize_filename(name, max_len=150):
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = re.sub(r"\s+", " ", name).strip(". ")
    return name[:max_len].rstrip(". ")


def get_json(session, url, params=None):
    for attempt in range(4):
        resp = session.get(url, params=params, timeout=60)
        if resp.status_code == 429:  # rate limited — back off and retry
            wait = 5 * (attempt + 1)
            print(f"  Rate limited, waiting {wait}s...")
            time.sleep(wait)
            continue
        resp.raise_for_status()
        return resp.json()
    resp.raise_for_status()


def fetch_conference(session, conf_id):
    data = get_json(session, f"{API}/conferences/{conf_id}")
    meta = data.get("metadata", {})
    title = (meta.get("titles") or [{}])[0].get("title", "")
    acronym = (meta.get("acronyms") or [""])[0]
    return title, acronym, meta.get("number_of_contributions")


def fetch_papers(session, conf_id):
    """Yield every literature record linked to the conference, following pagination."""
    params = {
        "q": f"publication_info.conference_record.$ref:{conf_id}",
        "size": 100,
        "sort": "mostrecent",
        "fields": "titles,authors.full_name,documents,arxiv_eprints,dois,"
                  "publication_info,control_number",
    }
    url = f"{API}/literature"
    while url:
        data = get_json(session, url, params)
        params = None  # the "next" link already contains the query
        hits = data.get("hits", {})
        for hit in hits.get("hits", []):
            yield hit.get("metadata", {})
        url = data.get("links", {}).get("next")
        time.sleep(INSPIRE_DELAY)


def pick_title(meta):
    titles = meta.get("titles") or []
    return titles[0].get("title", "").strip() if titles else ""


def pick_journal(meta):
    """Return the journal entry of publication_info (not the conference entry)."""
    for info in meta.get("publication_info") or []:
        if info.get("journal_title"):
            return info
    return {}


def pdf_sources(meta):
    """Return candidate PDF URLs in order of preference, tagged by source."""
    sources = []
    for doc in meta.get("documents") or []:
        if doc.get("hidden"):
            continue
        if doc.get("url"):
            sources.append(("inspire", doc["url"]))
    for eprint in meta.get("arxiv_eprints") or []:
        if eprint.get("value"):
            sources.append(("arxiv", f"https://arxiv.org/pdf/{eprint['value']}"))
    return sources


def download_pdf(session, url, local_path, source):
    """Download url to local_path. Returns True if the file now exists as a PDF."""
    if os.path.exists(local_path) and os.path.getsize(local_path) > 0:
        print(f"    Already exists: {os.path.basename(local_path)}")
        return True

    print(f"    Downloading ({source}): {url}")
    try:
        resp = session.get(url, timeout=120, stream=True, allow_redirects=True)
        resp.raise_for_status()
        first_chunk = b""
        tmp_path = local_path + ".part"
        with open(tmp_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                if not first_chunk:
                    first_chunk = chunk
                f.write(chunk)
        # Some sources return an HTML login/landing page instead of the PDF
        if not first_chunk.startswith(b"%PDF"):
            os.remove(tmp_path)
            print("    Not a PDF (probably restricted/paywalled), skipping this source")
            return False
        os.replace(tmp_path, local_path)
        print(f"    Saved: {os.path.basename(local_path)}")
        return True
    except Exception as e:
        print(f"    Failed: {e}")
        if os.path.exists(local_path + ".part"):
            os.remove(local_path + ".part")
        return False
    finally:
        time.sleep(ARXIV_DELAY if source == "arxiv" else INSPIRE_DELAY)


def scrape(conf_id, pdf_dir, output):
    os.makedirs(pdf_dir, exist_ok=True)
    session = requests.Session()
    session.headers.update(HEADERS)

    title, acronym, n_contrib = fetch_conference(session, conf_id)
    print(f"Conference: {title} ({acronym}), {n_contrib} contributions listed")

    rows = []
    no_pdf = []
    for i, meta in enumerate(fetch_papers(session, conf_id), start=1):
        recid = str(meta.get("control_number", ""))
        paper_title = pick_title(meta)
        print(f"\n[{i}] {paper_title[:80]}")

        journal = pick_journal(meta)
        dois = meta.get("dois") or []
        arxivs = meta.get("arxiv_eprints") or []

        # INSPIRE stores files with meaningless names ("document"), so name them
        # by record ID + title for easy browsing.
        filename = sanitize_filename(f"{recid} - {paper_title}") + ".pdf"
        local_path = os.path.join(pdf_dir, filename)

        saved = False
        for source, url in pdf_sources(meta):
            if download_pdf(session, url, local_path, source):
                saved = True
                break
        if not saved:
            print("    No downloadable PDF found")
            no_pdf.append(paper_title)

        rows.append([
            paper_title,
            "; ".join(a.get("full_name", "") for a in meta.get("authors") or []),
            journal.get("journal_title", ""),
            journal.get("journal_volume", ""),
            journal.get("artid", "") or journal.get("page_start", ""),
            journal.get("year", ""),
            dois[0].get("value", "") if dois else "",
            arxivs[0].get("value", "") if arxivs else "",
            recid,
            local_path if saved else "",
        ])

    with open(output, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(CSV_HEADERS)
        writer.writerows(rows)

    print(f"\n{'=' * 60}")
    print(f"Papers found:     {len(rows)}")
    print(f"PDFs downloaded:  {len(rows) - len(no_pdf)}")
    print(f"Without a PDF:    {len(no_pdf)}")
    print(f"CSV file:         {output}")
    print(f"PDF directory:    {pdf_dir}")
    if no_pdf:
        print("\nPapers with no downloadable PDF:")
        for t in no_pdf:
            print(f"  - {t}")


def main():
    parser = argparse.ArgumentParser(description="Download all PDFs for an INSPIRE-HEP conference")
    parser.add_argument("conference", nargs="?", default="2639288",
                        help="Conference ID or URL (default: 2639288, TWEPP23)")
    parser.add_argument("--pdf-dir", default=None, help="Folder for PDFs (default: inspire_<id>_pdfs)")
    parser.add_argument("--output", default=None, help="CSV file (default: inspire_<id>.csv)")
    args = parser.parse_args()

    conf_id = parse_conference_id(args.conference)
    scrape(
        conf_id,
        args.pdf_dir or f"inspire_{conf_id}_pdfs",
        args.output or f"inspire_{conf_id}.csv",
    )


if __name__ == "__main__":
    main()
