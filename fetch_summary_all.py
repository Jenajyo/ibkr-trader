import subprocess
import time
import requests
import os
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By

# === CONFIG ===
chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
user_data_dir = r"C:\Users\jyoti\AppData\Local\Google\Chrome\User Data"
remote_debugging_port = "9222"
lookback_hours = 2  # Change to 4, 5, or 12

now = datetime.now()
timestamp_str = now.strftime("%Y-%m-%d_%H-%M-%S")
output_dir = f"summaries_{timestamp_str}_{lookback_hours}h"
os.makedirs(output_dir, exist_ok=True)

# === FUNCTIONS ===
def kill_chrome():
    print("\U0001F534 Killing any existing Chrome processes...")
    os.system("taskkill /f /im chrome.exe >nul 2>&1")
    os.system("taskkill /f /im chromedriver.exe >nul 2>&1")

def is_debugger_running():
    try:
        res = requests.get(f"http://127.0.0.1:{remote_debugging_port}/json")
        return res.status_code == 200
    except:
        return False

def sanitize_filename(title):
    return re.sub(r'[\\/*?:"<>|]', "", title)

# === LAUNCH CHROME ===
kill_chrome()
if not is_debugger_running():
    subprocess.Popen([
        chrome_path,
        f'--remote-debugging-port={remote_debugging_port}',
        f'--user-data-dir={user_data_dir}',
        "--profile-directory=Default",
        "--no-first-run",
        "--no-default-browser-check",
        "--remote-allow-origins=*"
    ])
    print("\u23f3 Chrome launched. Waiting for debugger port to open...")
    for _ in range(30):
        if is_debugger_running():
            print("\u2705 Chrome debugger detected!")
            break
        time.sleep(1)
    else:
        raise Exception("\u274c Chrome debugger port did not open. Exiting...")

# === CONNECT DRIVER ===
options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", f"127.0.0.1:{remote_debugging_port}")
options.add_argument("--remote-allow-origins=*")
driver = webdriver.Chrome(options=options)

# === FETCH VIDEO LINKS ===
print("\U0001f4c5 Opening YouTube subscriptions feed...")
driver.get("https://www.youtube.com/feed/subscriptions")
time.sleep(5)

print("\U0001f501 Scrolling to collect video links...")
video_links = set()
keywords = ['seconds ago', 'minutes ago', 'hour ago', 'hours ago']
for _ in range(12):
    driver.execute_script("window.scrollBy(0, 2000);")
    time.sleep(2)
    items = driver.find_elements(By.CSS_SELECTOR, 'ytd-rich-item-renderer')
    for item in items:
        try:
            time_elements = item.find_elements(By.CSS_SELECTOR, '#metadata-line span')
            for time_element in time_elements:
                time_text = time_element.text.lower()
                if any(kw in time_text for kw in keywords):
                    digits = ''.join(c for c in time_text if c.isdigit())
                    age_minutes = 0
                    if 'minute' in time_text: age_minutes = int(digits)
                    elif 'second' in time_text: age_minutes = 0
                    elif 'hour' in time_text: age_minutes = int(digits) * 60
                    else: continue
                    if age_minutes <= lookback_hours * 60:
                        link_element = item.find_element(By.CSS_SELECTOR, 'a#thumbnail')
                        href = link_element.get_attribute('href')
                        if href and '/watch' in href:
                            video_links.add(href)
        except:
            continue

print(f"\u2705 Found {len(video_links)} videos in last {lookback_hours} hours.")

# === PROCESS VIDEOS ===
combined_summary = ""
for idx, link in enumerate(video_links):
    print(f"\n\U0001f3af Opening video {idx + 1}/{len(video_links)}: {link}")
    driver.get(link)
    time.sleep(6)

    input("\u23F3 Click 'Transcript & Summary', then press ENTER...\n")

    try:
        print("🔍 Extracting summary via visible widget...")

        # Fallback title from browser tab
        try:
            title = driver.execute_script("return document.title")
            if title.endswith(" - YouTube"):
                title = title.rsplit(" - YouTube", 1)[0].strip()
        except:
            title = "Untitled"

        filename_title = sanitize_filename(title)

        containers = driver.find_elements(By.CSS_SELECTOR, '[class*="glasp"]')
        found = False
        for container in containers:
            text = container.text.strip()
            if text and len(text) > 20:
                filepath = os.path.join(output_dir, f"{filename_title}.txt")
                summary_with_title = f"Title: {title}\nLink: {link}\n\n{text}"
                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(summary_with_title)
                combined_summary += f"\n{'='*80}\n{summary_with_title}\n"
                print(f"✅ Saved summary to {filepath}")
                found = True
                break
        if not found:
            print("⚠️ Summary block not found.")
    except Exception as e:
        print(f"❌ Error while extracting summary: {e}")

# === FINAL SUMMARY ===
if combined_summary.strip():
    final_file = os.path.join(output_dir, "combined_summaries.txt")
    with open(final_file, "w", encoding="utf-8") as f:
        f.write(combined_summary.strip())
    print(f"\n📍 Combined summary saved to: {final_file}")

print("\n✅ All done!")
driver.quit()
