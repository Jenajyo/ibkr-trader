from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import os

def get_glasp_summary_from_youtube(youtube_link):
    # Connect to already running Chrome with user profile
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    try:
        # Open new tab and navigate to the YouTube link
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])
        driver.get(youtube_link)

        print("✅ Opened YouTube video")

        # Wait for Glasp panel to load
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class,'glasp-sidebar') or contains(text(),'Transcript & Summary')]")
        ))
        print("✅ Glasp panel detected")

        # Click the "Summary" tab
        try:
            summary_tab = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//div[@role='tab' and contains(.,'Summary')]")
            ))
            driver.execute_script("arguments[0].click();", summary_tab)
            print("✅ Clicked on Summary tab")
        except Exception as e:
            print(f"⚠️ Couldn't find or click Summary tab: {e}")

        time.sleep(2)  # Give time for summary content to load

        # Now scrape the Summary
        summary_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'summary')]")
        ))
        summary_text = summary_element.text

        if summary_text.strip():
            print("✅ Summary fetched successfully")

            # Save to file
            with open("glasp_summary.txt", "w", encoding="utf-8") as f:
                f.write(summary_text)
            print("✅ Saved to glasp_summary.txt")
        else:
            print("❌ Empty summary fetched")

    except Exception as e:
        print(f"❌ Error: {e}")

    finally:
        driver.close()

if __name__ == "__main__":
    youtube_link = input("https://www.youtube.com/watch?v=0snE3qiiZcs&t=3s")
    get_glasp_summary_from_youtube(youtube_link)
