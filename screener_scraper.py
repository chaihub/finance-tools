#!/usr/bin/env python3
"""
Screener.in Data Scraper
Automated script to download paginated financial data into a CSV file.

SETUP INSTRUCTIONS:
1. Run this script: python screener_scraper.py
2. Firefox will open
3. Navigate to your Screener page and log in and setup the data that you wish to extract
3a. Make sure you can see the data table with rows and the "Next >" button
4. Press Enter to start scraping

The script will detect the page URL and use it in the output file.
"""

import asyncio
import csv
from datetime import datetime
from pathlib import Path
from playwright.async_api import async_playwright
import pytz

# Configuration
CSV_FILE_PATH = "data/screener_data.csv"

async def scrape_screener_data():
    """Scrape all paginated data from Screener.in and save to CSV."""
    
    all_rows = []
    headers = None
    page_count = 0
    page_url = None
    
    async with async_playwright() as p:
        # Connect to existing Firefox instance or launch new one
        browser = await p.firefox.launch(headless=False)
        page = await browser.new_page()
        
        print(f"\n{'='*70}")
        print("SCREENER.IN DATA SCRAPER")
        print(f"{'='*70}")
        print(f"\nIMPORTANT: If a new Firefox window opened, please:")
        print("1. Navigate to your Screener page")
        print("2. Wait for the data table to fully load")
        print("3. Return to this terminal and press Enter when ready")
        print(f"{'='*70}\n")
        
        input("Press Enter once you have the Screener page loaded in Firefox...")
        
        # Detect the current page URL
        page_url = page.url
        print(f"\n✓ Detected URL: {page_url}\n")
        
        while True:
            page_count += 1
            print(f"\n📄 Scraping page {page_count}...")
            
            # Wait for table with extended timeout
            try:
                await page.wait_for_selector("table", timeout=2000)
                print("   ✓ Table element found")
            except Exception as e:
                print(f"⚠️  Timeout waiting for table element")
                # Debug: Save screenshot and HTML for inspection
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                await page.screenshot(path=f"debug_screenshot_{timestamp}.png")
                html_content = await page.content()
                with open(f"debug_html_{timestamp}.txt", "w", encoding="utf-8") as f:
                    f.write(html_content)
                print(f"   Debug files saved: debug_screenshot_{timestamp}.png, debug_html_{timestamp}.txt")
                break
            
            # Check if table exists
            table_exists = await page.query_selector("table")
            if not table_exists:
                print("❌ No table found on page. Pagination may have ended.")
                break
            
            # Extract headers from first page
            if headers is None:
                # Try multiple header selectors for robustness
                header_cells = await page.query_selector_all("table thead th")
                if not header_cells:
                    header_cells = await page.query_selector_all("table tr:first-child th")
                if not header_cells:
                    header_cells = await page.query_selector_all("table tr:first-child td")
                
                headers = [await cell.text_content() for cell in header_cells]
                headers = [h.strip() for h in headers if h.strip()]
                
                if not headers:
                    print("❌ No headers found. Trying alternative extraction...")
                    # Fallback 1: Extract from first row of tds
                    first_row = await page.query_selector("table tr")
                    if first_row:
                        first_cells = await first_row.query_selector_all("td, th")
                        headers = [await cell.text_content() for cell in first_cells]
                        headers = [h.strip() for h in headers if h.strip()]
                    
                    if not headers:
                        print("❌ Still no headers found. Unable to parse table structure.")
                        await browser.close()
                        return
                print(f"   ✓ Found {len(headers)} columns")
            
            # Extract row data - try multiple selectors for robustness
            rows = await page.query_selector_all("table tbody tr")
            if not rows:
                # Fallback: try without tbody
                rows = await page.query_selector_all("table tr:not(:first-child)")
            
            rows_extracted = 0
            for row in rows:
                cells = await row.query_selector_all("td")
                if not cells:
                    cells = await row.query_selector_all("th")  # In case it's a header row
                row_data = [await cell.text_content() for cell in cells]
                row_data = [cell.strip() for cell in row_data]
                if row_data and len(row_data) > 1:  # Skip empty rows
                    all_rows.append(row_data)
                    rows_extracted += 1
            
            print(f"   ✓ Extracted {rows_extracted} rows")
            
            # Check if Next button exists and is enabled
            try:
                # Try multiple selectors for the Next button
                next_button = await page.query_selector("a:has-text('Next')")
                if not next_button:
                    next_button = await page.query_selector("a:contains('Next')")
                if not next_button:
                    # Look for any link with text containing "Next"
                    all_links = await page.query_selector_all("a")
                    for link in all_links:
                        link_text = await link.text_content()
                        if "Next" in link_text:
                            next_button = link
                            break
                
                if next_button:
                    # Check if button's parent li is disabled
                    parent_li = await next_button.evaluate("el => el.closest('li')")
                    if parent_li:
                        is_disabled = await page.evaluate("(li) => li.classList.contains('disabled')", parent_li)
                        if is_disabled:
                            print("\n✅ Reached last page. Pagination complete.")
                            break
                    
                    print("   Clicking Next button...")
                    await next_button.click()
                    # Wait for new content to load
                    await asyncio.sleep(2)
                    await page.wait_for_load_state("networkidle")
                    print("   ✓ Page loaded")
                else:
                    print("\n✅ No more 'Next' button found. Pagination complete.")
                    break
            except Exception as e:
                print(f"❌ Pagination error: {e}")
                print("   Stopping scrape.")
                break
        
        await browser.close()
    
    # Write to CSV
    csv_path = Path(CSV_FILE_PATH)
    file_exists = csv_path.exists()
    
    with open(csv_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        
        # Add metadata on first line only if file is new
        if not file_exists:
            tz = pytz.timezone('Asia/Kolkata')
            timestamp = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S %Z")
            writer.writerow([f"Source: {page_url}"])
            writer.writerow([f"Extracted: {timestamp}"])
            writer.writerow([])  # Blank line for clarity
        
        # Write headers only on new file
        if not file_exists and headers:
            writer.writerow(headers)
        
        # Write all data rows
        if all_rows:
            writer.writerows(all_rows)
    
    print(f"\n{'='*70}")
    print("✓ SCRAPING COMPLETE")
    print(f"{'='*70}")
    print(f"Total pages scraped: {page_count}")
    print(f"Total rows extracted: {len(all_rows)}")
    print(f"CSV file: {csv_path.absolute()}")
    print(f"{'='*70}\n")

if __name__ == "__main__":
    asyncio.run(scrape_screener_data())
