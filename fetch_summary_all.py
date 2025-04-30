import subprocess
import time
import requests
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
user_data_dir = r"C:\Users\jyoti\AppData\Local\Google\Chrome\User Data"
remote_debugging_port = "9222"

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

options = webdriver.ChromeOptions()
options.add_experimental_option("debuggerAddress", f"127.0.0.1:{remote_debugging_port}")
options.add_argument("--remote-allow-origins=*")

driver = webdriver.Chrome(options=options)

# Step 1: Open subscriptions
print("\U0001f4c5 Opening YouTube subscriptions feed...")
driver.get("https://www.youtube.com/feed/subscriptions")
time.sleep(5)

# Step 2: Scroll and gather video links
print("\U0001f501 Scrolling to collect video links from today...")
video_links = set()
for _ in range(10):
    driver.execute_script("window.scrollBy(0, 2000);")
    time.sleep(2)
    items = driver.find_elements(By.CSS_SELECTOR, 'ytd-rich-item-renderer')
    for item in items:
        try:
            time_elements = item.find_elements(By.CSS_SELECTOR, '#metadata-line span')
            for time_element in time_elements:
                time_text = time_element.text.lower()
                if any(kw in time_text for kw in ['seconds ago', 'minutes ago', 'hours ago', 'streamed', 'just now']):
                    link_element = item.find_element(By.CSS_SELECTOR, 'a#thumbnail')
                    href = link_element.get_attribute('href')
                    if href and '/watch' in href:
                        video_links.add(href)
                        break
        except:
            continue

print(f"\u2705 Found {len(video_links)} videos from today.")

# Step 3: Process each video one by one
for idx, link in enumerate(video_links):
    print(f"\n\U0001f3af Opening video {idx + 1}/{len(video_links)}: {link}")
    driver.get(link)
    time.sleep(5)

    driver.execute_script("window.scrollBy(0, 800);")
    time.sleep(2)

    input("\n⏳ Click 'Transcript & Summary' for this video, then press ENTER to continue...\n")

    try:
        print("🔍 Searching for summary block...")
        containers = driver.find_elements(By.CSS_SELECTOR, '[class*="glasp"]')
        found = False
        for container in containers:
            text = container.text.strip()
            if text.startswith("Transcript & Summary"):
                parts = text.split("Transcript & Summary", 1)
                summary = parts[1].strip()
                with open(f"glasp_summary_{idx + 1}.txt", "w", encoding="utf-8") as f:
                    f.write(summary)
                print(f"✅ Saved summary to glasp_summary_{idx + 1}.txt")
                found = True
                break
        if not found:
            print("⚠️ Summary block not found.")
    except Exception as e:
        print(f"\u274c Error while extracting summary: {e}")

driver.quit()
