#!/usr/bin/env python3
"""
NASA Radiation Database Scraper with PDF Download
Scrapes all pages, extracts PDF links from popups, downloads PDFs, and creates CSV with local paths

Usage:
    python nasa_scraper_with_pdfs.py [--pdf-dir DIR] [--output FILE] [--max-pages N] [--no-headless]

Requirements:
    pip install selenium webdriver-manager requests
"""

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
import csv
import time
import sys
import os
import requests
import re
from urllib.parse import urljoin, urlparse


def setup_driver(headless=True):
    """Initialize Chrome driver with appropriate options"""
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    
    driver = None
    
    # Try webdriver-manager first
    try:
        from selenium.webdriver.chrome.service import Service as ChromeService
        from webdriver_manager.chrome import ChromeDriverManager
        print("Using webdriver-manager to get ChromeDriver...")
        service = ChromeService(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        print(f"webdriver-manager failed: {e}")
        print("Trying system chromedriver...")
        try:
            driver = webdriver.Chrome(options=chrome_options)
        except Exception as e2:
            print(f"System chromedriver failed: {e2}")
            print("\nPlease install Chrome and chromedriver, or run:")
            print("  pip install webdriver-manager")
            sys.exit(1)
    
    return driver


def sanitize_filename(filename):
    """Sanitize filename to be safe for filesystem"""
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    filename = filename.strip('. ')
    if len(filename) > 200:
        filename = filename[:200]
    return filename


def download_pdf(url, pdf_dir, base_url):
    """Download a PDF file and return the local path"""
    full_url = urljoin(base_url, url)
    
    try:
        parsed = urlparse(full_url)
        original_filename = os.path.basename(parsed.path)
        if not original_filename.lower().endswith('.pdf'):
            original_filename += '.pdf'
        
        safe_filename = sanitize_filename(original_filename)
        local_path = os.path.join(pdf_dir, safe_filename)
        
        # Check if file already exists (same filename = assume same PDF)
        if os.path.exists(local_path):
            print(f"    Already exists: {safe_filename}")
            return local_path
        
        print(f"    Downloading: {full_url}")
        response = requests.get(full_url, timeout=60, stream=True)
        response.raise_for_status()
        
        with open(local_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
        
        print(f"    Saved: {safe_filename}")
        return local_path
        
    except Exception as e:
        print(f"    Failed to download {full_url}: {e}")
        return None


def extract_pdf_links_from_popup(driver, wait):
    """Extract PDF links from the popup dialog (jqGrid ViewRowData dialog)"""
    pdf_links = []
    
    try:
        # jqGrid uses ViewRowData dialog - wait for it to appear
        # The dialog typically has id like "ViewModlist" or contains "viewmod"
        popup_selectors = [
            "#ViewModlist",
            "#viewmod",
            ".ui-jqdialog",
            ".ui-dialog",
            "[role='dialog']",
            ".view-modal",
        ]
        
        popup = None
        for selector in popup_selectors:
            try:
                elements = driver.find_elements(By.CSS_SELECTOR, selector)
                for elem in elements:
                    if elem.is_displayed():
                        popup = elem
                        break
                if popup:
                    break
            except:
                continue
        
        # If still no popup, wait briefly and try again
        if not popup:
            time.sleep(1)
            for selector in popup_selectors:
                try:
                    elements = driver.find_elements(By.CSS_SELECTOR, selector)
                    for elem in elements:
                        if elem.is_displayed():
                            popup = elem
                            break
                    if popup:
                        break
                except:
                    continue
        
        if popup:
            # Look for PDF links within the popup
            link_elements = popup.find_elements(By.TAG_NAME, "a")
            
            for link in link_elements:
                href = link.get_attribute("href")
                if href:
                    href_lower = href.lower()
                    text_lower = (link.text or "").lower()
                    # Check for PDF in href or link text
                    if '.pdf' in href_lower or 'pdf' in text_lower:
                        pdf_links.append(href)
        
        # Also check the entire page for newly appeared PDF links
        # jqGrid might load content dynamically
        if not pdf_links:
            # Look for PDF links that might be in a table cell or form element
            all_pdf_links = driver.find_elements(By.XPATH, 
                "//a[contains(@href, '.pdf') or contains(@href, '.PDF')]")
            for link in all_pdf_links:
                href = link.get_attribute("href")
                if href and link.is_displayed():
                    pdf_links.append(href)
        
        # Check for links in the Files column content that might be visible
        # The popup might show the "Files" field value which contains links
        files_cells = driver.find_elements(By.XPATH, 
            "//td[contains(@class, 'DataTD') or contains(@class, 'CaptionTD')]//a")
        for cell in files_cells:
            href = cell.get_attribute("href")
            if href and '.pdf' in href.lower():
                pdf_links.append(href)
        
    except Exception as e:
        print(f"    Error extracting PDF links: {e}")
    
    # Remove duplicates while preserving order
    seen = set()
    unique_links = []
    for link in pdf_links:
        if link not in seen:
            seen.add(link)
            unique_links.append(link)
    
    return unique_links


def close_popup(driver):
    """Close any open popup dialog"""
    close_selectors = [
        ".ui-jqdialog-titlebar-close",
        ".ui-dialog-titlebar-close",
        "span.ui-icon-closethick",
        "button.close",
        "[aria-label='close']",
        "[aria-label='Close']",
        ".modal-header .close"
    ]
    
    for selector in close_selectors:
        try:
            close_btns = driver.find_elements(By.CSS_SELECTOR, selector)
            for close_btn in close_btns:
                if close_btn.is_displayed():
                    try:
                        close_btn.click()
                    except:
                        driver.execute_script("arguments[0].click();", close_btn)
                    time.sleep(0.5)
                    return True
        except:
            continue
    
    # Try clicking the close icon's parent element
    try:
        close_icons = driver.find_elements(By.CSS_SELECTOR, ".ui-icon-closethick")
        for icon in close_icons:
            parent = icon.find_element(By.XPATH, "..")
            if parent.is_displayed():
                try:
                    parent.click()
                except:
                    driver.execute_script("arguments[0].click();", parent)
                time.sleep(0.5)
                return True
    except:
        pass
    
    # Try pressing Escape
    try:
        from selenium.webdriver.common.keys import Keys
        driver.find_element(By.TAG_NAME, "body").send_keys(Keys.ESCAPE)
        time.sleep(0.5)
        return True
    except:
        pass
    
    return False


def scrape_with_pdfs(url="https://nepp.nasa.gov/radhome/RadDatabase/RadDataBase.html", pdf_dir="pdfs", headless=True, max_pages=200):
    """Scrape all pages, extract PDF links, and download PDFs"""
    print(f"Launching browser and loading {url}...")
    
    os.makedirs(pdf_dir, exist_ok=True)
    
    driver = setup_driver(headless=headless)
    base_url = url.rsplit('/', 1)[0] + '/'
    
    try:
        driver.get(url)
        
        print("Waiting for table to load...")
        wait = WebDriverWait(driver, 20)
        table = wait.until(EC.presence_of_element_located((By.TAG_NAME, "table")))
        time.sleep(3)
        
        headers = []
        header_elements = driver.find_elements(By.CSS_SELECTOR, "table thead th")
        if not header_elements:
            header_elements = driver.find_elements(By.CSS_SELECTOR, "table tr:first-child th")
        if not header_elements:
            header_elements = driver.find_elements(By.CSS_SELECTOR, "table tr:first-child td")
        
        headers = [elem.text.strip() for elem in header_elements if elem.text.strip()]
        print(f"Found {len(headers)} columns: {headers}")
        
        # Keep only the meaningful columns (first 7) plus our PDF paths column
        # The actual data columns are: Part Number, Function, Manufacturer, Date(s), File(s), Test Type, Category
        meaningful_headers = ['Part Number', 'Function', 'Manufacturer', 'Date(s)', 'File(s)', 'Test Type', 'Category', 'PDF_Local_Paths']
        
        if len(headers) < 2:
            print("ERROR: No headers found!")
            return [], []
        
        all_data_rows = []
        current_page = 1
        row_id_tracker = set()
        
        print(f"\nStarting to scrape pages with PDF extraction (max {max_pages} pages)...")
        
        while current_page <= max_pages:
            print(f"\n{'='*60}")
            print(f"Processing page {current_page}...")
            time.sleep(2)
            
            # Get the number of rows on this page first
            row_elements = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            if not row_elements:
                row_elements = driver.find_elements(By.CSS_SELECTOR, "table tr")[1:]
            
            num_rows = len(row_elements)
            print(f"  Found {num_rows} rows on this page")
            
            page_rows = 0
            
            # Use index-based iteration to avoid stale element issues
            for row_idx in range(num_rows):
                # Re-fetch the row element each time (avoids stale reference)
                try:
                    row_elements = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                    if not row_elements:
                        row_elements = driver.find_elements(By.CSS_SELECTOR, "table tr")[1:]
                    
                    if row_idx >= len(row_elements):
                        break
                    
                    row = row_elements[row_idx]
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if not cells:
                        continue
                    
                    # Get raw cell data and clean it up
                    row_data = []
                    for cell in cells:
                        text = cell.text.strip()
                        # Replace newlines with semicolons for cleaner CSV
                        text = re.sub(r'\n+', '; ', text)
                        # Clean up multiple spaces
                        text = re.sub(r'\s+', ' ', text)
                        row_data.append(text)
                    
                    if not any(row_data):
                        continue
                    
                    # Skip header/filter rows and junk data
                    first_cell = row_data[0].upper() if row_data else ''
                    if 'CLEAR SEARCH' in first_cell or 'RESET TABLE' in first_cell:
                        continue
                    if 'PAGE' in first_cell and 'OF' in first_cell:
                        continue
                    
                    # Skip rows that are mostly empty or just contain 'x' placeholders
                    non_empty_cells = [c for c in row_data[:7] if c and c.lower() != 'x']
                    if len(non_empty_cells) < 2:
                        continue
                    
                    row_id = "|".join(row_data[:3])
                    if row_id in row_id_tracker:
                        continue
                    row_id_tracker.add(row_id)
                    
                    part_num = row_data[0] if row_data else 'N/A'
                    print(f"  Row {page_rows + 1}: {part_num[:50]}...")
                    
                    pdf_links = []
                    local_paths = []
                    
                    try:
                        # Re-fetch the row again before clicking (it may have gone stale)
                        row_elements = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                        if not row_elements:
                            row_elements = driver.find_elements(By.CSS_SELECTOR, "table tr")[1:]
                        
                        if row_idx < len(row_elements):
                            row = row_elements[row_idx]
                            
                            actions = ActionChains(driver)
                            driver.execute_script("arguments[0].scrollIntoView(true);", row)
                            time.sleep(0.3)
                            actions.double_click(row).perform()
                            time.sleep(1.5)
                            
                            pdf_links = extract_pdf_links_from_popup(driver, wait)
                            
                            if pdf_links:
                                print(f"    Found {len(pdf_links)} PDF link(s)")
                                for pdf_url in pdf_links:
                                    local_path = download_pdf(pdf_url, pdf_dir, base_url)
                                    if local_path:
                                        local_paths.append(local_path)
                            else:
                                print(f"    No PDFs found")
                            
                            close_popup(driver)
                            time.sleep(0.5)
                        
                    except Exception as e:
                        print(f"    Error processing row popup: {e}")
                        close_popup(driver)
                        time.sleep(0.5)
                    
                    # Keep only the first 7 meaningful columns + PDF paths
                    clean_row = row_data[:7]
                    # Pad if needed
                    while len(clean_row) < 7:
                        clean_row.append("")
                    # Add PDF local paths
                    clean_row.append("; ".join(local_paths) if local_paths else "")
                    
                    all_data_rows.append(clean_row)
                    page_rows += 1
                    
                except Exception as e:
                    print(f"    Error on row {row_idx}: {e}")
                    continue
            
            print(f"  Extracted {page_rows} rows from page {current_page}")
            print(f"  Total rows so far: {len(all_data_rows)}")
            
            next_button = None
            selectors = [
                "span.ui-icon.ui-icon-seek-next",
                "button span.ui-icon-seek-next",
                "a span.ui-icon-seek-next",
                ".ui-icon-seek-next",
            ]
            
            for selector in selectors:
                try:
                    icon_element = driver.find_element(By.CSS_SELECTOR, selector)
                    next_button = icon_element.find_element(By.XPATH, "..")
                    if next_button.is_displayed() and next_button.is_enabled():
                        break
                    else:
                        next_button = None
                except:
                    continue
            
            if next_button:
                driver.execute_script("arguments[0].scrollIntoView(true);", next_button)
                time.sleep(0.5)
                
                try:
                    next_button.click()
                except:
                    driver.execute_script("arguments[0].click();", next_button)
                
                current_page += 1
                time.sleep(1.5)
            else:
                print(f"\n  No more pages. Stopping at page {current_page}")
                break
        
        print(f"\n{'='*60}")
        print(f"✓ Extraction complete!")
        print(f"  Total pages processed: {current_page}")
        print(f"  Total rows extracted: {len(all_data_rows)}")
        
        pdf_count = sum(1 for row in all_data_rows if len(row) > 7 and row[7])
        print(f"  Rows with PDFs: {pdf_count}")
        
        return meaningful_headers, all_data_rows
    
    finally:
        if driver:
            driver.quit()


def save_to_csv(headers, data_rows, output_file='nasa_radiation_data_with_pdfs.csv'):
    """Save the scraped data to a CSV file"""
    if not data_rows:
        print("No data to save!")
        return
    
    print(f"\nSaving to {output_file}...")
    
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        writer.writerows(data_rows)
    
    print(f"✓ Successfully saved {len(data_rows)} rows to {output_file}")


def main():
    import argparse
    parser = argparse.ArgumentParser(description='NASA Radiation Database Scraper with PDF Download')
    parser.add_argument('--pdf-dir', default='pdfs', help='Directory to save PDFs (default: pdfs)')
    parser.add_argument('--output', default='nasa_radiation_data_with_pdfs.csv', help='Output CSV file')
    parser.add_argument('--max-pages', type=int, default=200, help='Maximum pages to scrape')
    parser.add_argument('--no-headless', action='store_true', help='Run browser in visible mode (useful for debugging)')
    args = parser.parse_args()
    
    try:
        headers, data_rows = scrape_with_pdfs(
            pdf_dir=args.pdf_dir,
            headless=not args.no_headless,
            max_pages=args.max_pages
        )
        save_to_csv(headers, data_rows, args.output)
        
        print(f"\n{'='*60}")
        print("SUMMARY")
        print(f"{'='*60}")
        print(f"CSV file: {args.output}")
        print(f"PDF directory: {args.pdf_dir}")
        
        if os.path.exists(args.pdf_dir):
            pdf_files = [f for f in os.listdir(args.pdf_dir) if f.endswith('.pdf')]
            print(f"Total PDFs downloaded: {len(pdf_files)}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()