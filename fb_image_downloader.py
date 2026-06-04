from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import os
import re
import time

import requests

# ---------------- CONFIG ----------------
START_URL = "https://www.facebook.com/photo/?fbid=855794077291877&set=g.1668337183604527"
SAVE_DIR = r"C:\Users\safid\Downloads\Memes 3"
STATE_FILE = "scraper_state.json"

# ---------------- SETUP ----------------
os.makedirs(SAVE_DIR, exist_ok=True)


# -------- Auto-detect last saved number from folder --------
def get_last_saved_number(folder):
    """Scan folder for files like 0001.jpg, 0002.jpg and return the highest number found."""
    max_num = 0
    pattern = re.compile(r"^(\d+)\.(jpg|jpeg|png|webp)$", re.IGNORECASE)
    if os.path.exists(folder):
        for fname in os.listdir(folder):
            match = pattern.match(fname)
            if match:
                num = int(match.group(1))
                if num > max_num:
                    max_num = num
    return max_num


auto_last = get_last_saved_number(SAVE_DIR)

# -------- Load or Reset State --------
state = {
    "start_url": START_URL,
    "last_url": START_URL,
    "last_number": auto_last,
}

if os.path.exists(STATE_FILE):
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            raw_state = f.read().strip()
        old_state = json.loads(raw_state) if raw_state else None
    except (json.JSONDecodeError, OSError) as e:
        old_state = None
        print(f"State file is empty or invalid ({e}). Starting fresh from folder scan: {auto_last}")

    if isinstance(old_state, dict) and old_state.get("start_url") == START_URL:
        # Use whichever is higher: state file or folder scan (folder is source of truth)
        state = old_state
        state["last_number"] = max(old_state.get("last_number", 0), auto_last)
        print(f"Resuming from image #{state['last_number']} (folder scan: {auto_last})")
    elif isinstance(old_state, dict):
        print(f"START_URL changed; starting fresh. Folder scan found last number: {auto_last}")
    else:
        print(f"No usable state found. Folder scan found last number: {auto_last}")
else:
    print(f"No state file found. Folder scan found last number: {auto_last}")

saved = state["last_number"]

# Extract starting FBID (more reliable loop detection)
START_FBID = START_URL.split("fbid=")[1].split("&")[0]


# ---------------- HELPERS ----------------
def is_video_page(driver):
    """Check if the current page/post contains a video instead of a photo."""
    try:
        videos = driver.find_elements(By.CSS_SELECTOR, "video")
        if videos:
            return True

        current_url = driver.current_url
        if "video" in current_url.lower():
            return True

        video_selectors = [
            "div[data-pagelet='VideoPlayer']",
            "div[class*='videoPlayer']",
            "div[aria-label='Video player']",
        ]
        for sel in video_selectors:
            if driver.find_elements(By.CSS_SELECTOR, sel):
                return True

    except Exception:
        pass
    return False


# ---------------- DRIVER ----------------
driver = webdriver.Chrome()
driver.get(state["last_url"])

input("Log in manually, then press Enter to start scraping...")

while True:
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "body"))
        )

        if is_video_page(driver):
            print(f"Skipping video at: {driver.current_url}")
        else:
            try:
                WebDriverWait(driver, 8).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "img[src*='scontent']"))
                )
            except Exception:
                print("No scontent images found on this page, skipping.")

            imgs = driver.find_elements(By.CSS_SELECTOR, "img[src*='scontent']")

            largest_img = None
            largest_area = 0

            for img in imgs:
                try:
                    width = driver.execute_script("return arguments[0].naturalWidth;", img)
                    height = driver.execute_script("return arguments[0].naturalHeight;", img)
                    area = width * height

                    if area > largest_area:
                        largest_area = area
                        largest_img = img
                except Exception:
                    continue

            if not largest_img or largest_area == 0:
                print("Could not find main photo, skipping this entry.")
            else:
                src = largest_img.get_attribute("src")

                saved += 1
                filename = os.path.join(SAVE_DIR, f"{saved:04d}.jpg")

                r = requests.get(src, timeout=15)
                with open(filename, "wb") as f:
                    f.write(r.content)

                print(f"[{saved}] saved: {filename}")

                current_url = driver.current_url
                state = {
                    "start_url": START_URL,
                    "last_url": current_url,
                    "last_number": saved,
                }

                with open(STATE_FILE, "w", encoding="utf-8") as f:
                    json.dump(state, f, indent=4)

        next_btn = None
        btn_selectors = [
            "div[aria-label='Next']",
            "a[aria-label='Next']",
            "div[aria-label='Next photo']",
        ]

        for sel in btn_selectors:
            try:
                next_btn = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, sel))
                )
                if next_btn:
                    break
            except Exception:
                continue

        if next_btn:
            driver.execute_script("arguments[0].click();", next_btn)
            time.sleep(2)

            new_url = driver.current_url

            if "fbid=" in new_url:
                current_fbid = new_url.split("fbid=")[1].split("&")[0]

                if current_fbid == START_FBID and saved > 0:
                    print("Reached starting photo again; stopping scraper.")
                    break
        else:
            print("No more next button; stopping.")
            break

    except Exception as e:
        print(f"Error: {e}")
        break

driver.quit()
print(f"Done. Total images saved: {saved}")