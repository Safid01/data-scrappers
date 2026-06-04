import os
import tempfile
import datetime
import re
import time
from urllib.parse import urljoin

import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


BASE_URL = "https://www.rokomari.com/reviews?type=1&ref=mr_p0&page={}"
MAX_RATINGS = 3993
MAX_REVIEWS = 2152

def extract_number(text):
    if not text:
        return 0
    m = re.search(r'(\d{1,3}(?:,\d{3})*|\d+)', text.replace(",", ""))
    return int(m.group(1)) if m else 0

def quality_label(num_ratings, num_reviews):
    s_r = (num_ratings / MAX_RATINGS) if MAX_RATINGS else 0
    s_v = (num_reviews / MAX_REVIEWS) if MAX_REVIEWS else 0
    final = (s_r + s_v) / 2.0
    if final < 0.33:
        return "Not Good"
    elif final < 0.66:
        return "Good"
    else:
        return "Best"

def parse_listing_page(driver, base_url="https://www.rokomari.com"):
    books = []
    seen = set()
    anchors = driver.find_elements(By.XPATH, "//a[contains(@href, '/book/')]")
    for a in anchors:
        try:
            href = a.get_attribute("href")
            title = a.text.strip()
            if not href or not title or href in seen:
                continue
            seen.add(href)

            parent = a.find_element(By.XPATH, "./ancestor::div[1]")
            text = parent.text

            try:
                author_el = parent.find_element(By.XPATH, ".//a[contains(@href,'/author/')]")
                author = author_el.text.strip()
            except:
                author = ""

            m_rat = re.search(r'(\d[\d,]*)\s*Ratings', text)
            num_ratings = int(m_rat.group(1).replace(",", "")) if m_rat else 0

            m_rev = re.search(r'(\d[\d,]*)\s*Reviews', text)
            num_reviews = int(m_rev.group(1).replace(",", "")) if m_rev else 0

            books.append({
                "book_title": title,
                "author": author,
                "num_ratings": num_ratings,
                "num_reviews": num_reviews,
                "quality": quality_label(num_ratings, num_reviews),
                "book_url": href,
            })
        except:
            continue
    return books

def _save_progress(all_books, out_file, page=None):
    """Write current progress to an Excel file atomically.

    Tries to write to a temporary file and then replace the target. On Windows
    this can fail if the target file is open in Excel; in that case we fall
    back to writing a timestamped copy so progress isn't lost.
    """
    df = pd.DataFrame(all_books.values())
    # write to a temp file in same directory to allow atomic replace
    out_dir = os.path.dirname(os.path.abspath(out_file)) or '.'
    tmp_fd, tmp_path = tempfile.mkstemp(suffix='.xlsx', dir=out_dir)
    os.close(tmp_fd)
    try:
        df.to_excel(tmp_path, index=False)
        try:
            # atomic replace (works on Windows Python 3.3+)
            os.replace(tmp_path, out_file)
            stamp = f" after page {page}" if page else ""
            print(f"[✓] Saved {len(df)} books → {out_file}{stamp}")
        except PermissionError:
            # target may be open in Excel; write a timestamped backup instead
            ts = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
            fallback = f"{os.path.splitext(out_file)[0]}_{ts}.xlsx"
            df.to_excel(fallback, index=False)
            print(f"[!] Could not overwrite {out_file} (file locked). Wrote fallback {fallback}")
            try:
                os.remove(tmp_path)
            except:
                pass
    except Exception as e:
        print(f"[!] Failed to save progress: {e}")
        try:
            os.remove(tmp_path)
        except:
            pass


def scrape_all_selenium(max_pages=2000, out_file="rokomari_books.xlsx", save_interval=10):
    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    driver = webdriver.Chrome(service=Service(), options=options)

    all_books = {}
    last_page = None
    interrupted = False
    try:
        for page in range(1, max_pages+1):
            last_page = page
            url = BASE_URL.format(page)
            print(f"[+] Page {page}: {url}")
            driver.get(url)

            # auto-accept cookie popup if exists
            try:
                els = driver.find_elements(By.XPATH, "//button[contains(.,'Accept') or contains(.,'সম্মতি') or contains(.,'agree')]")
                if els:
                    try:
                        els[0].click()
                        print("   [✓] Cookie consent accepted")
                    except:
                        pass
            except:
                pass

            time.sleep(2)
            books = parse_listing_page(driver)
            if not books:
                print("   [!] No books found => stopping")
                break

            for b in books:
                if b['book_url'] not in all_books:
                    all_books[b['book_url']] = b
            print(f"   Collected {len(books)}, total {len(all_books)}")

            # Periodically save progress so we don't lose everything on crash
            if save_interval and page % save_interval == 0:
                _save_progress(all_books, out_file, page=page)

    except KeyboardInterrupt:
        interrupted = True
        print("\n[!] Interrupted by user — saving progress before exit...")
    finally:
        try:
            driver.quit()
        except:
            pass
        # final save
        _save_progress(all_books, out_file, page=last_page)
        if interrupted:
            print("[✓] Progress saved. Exiting due to user interrupt.")

if __name__ == "__main__":
    scrape_all_selenium()