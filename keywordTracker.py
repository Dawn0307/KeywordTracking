import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def start_chrome_with_existing_profile():
    options = Options()
    user_data_dir = 'C:\\Users\\A\\AppData\\Local\\Google\\Chrome\\User Data'
    profile_dir = 'Default'

    options.add_argument(f"user-data-dir={user_data_dir}")
    options.add_argument(f"profile-directory={profile_dir}")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    
    return driver

driver = start_chrome_with_existing_profile()
url = 'https://members.helium10.com/keyword-tracker?accountId=1545862407'
driver.get(url)

try:
    # 等待加載
    rows = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tr.kt-orders-row"))
    )
    
    # 找到目標 data-key
    target_data_key = "1988788"
    target_row = None
    for row in rows:
        if row.get_attribute("data-key") == target_data_key:
            target_row = row
            break

    if target_row:
        # 提取數據
        product_name = target_row.find_element(By.CSS_SELECTOR, ".media-title").text
        asin = target_row.find_element(By.CSS_SELECTOR, ".media-asin .show-asin-variations").text
        tracked_keywords = target_row.find_element(By.CSS_SELECTOR, ".media-keywords-tracked span").text.split(":")[-1].strip()
        top_10_keywords = target_row.find_element(By.XPATH, ".//p[text()='Top 10']/preceding-sibling::span[1]").text
        top_50_keywords = target_row.find_element(By.XPATH, ".//p[text()='Top 50']/preceding-sibling::span[1]").text
        top_10_search_volume = target_row.find_element(By.XPATH, ".//div[@class='col-md-6 search_volume_keywords']//p[text()='Top 10']/preceding-sibling::span[1]").text
        top_50_search_volume = target_row.find_element(By.XPATH, ".//div[@class='col-md-6 search_volume_keywords']//p[text()='Top 50']/preceding-sibling::span[1]").text

        print("Product Name:", product_name)
        print("ASIN:", asin)
        print("Tracked Keywords:", tracked_keywords)
        print("Top 10 Search Volume:", top_10_search_volume)
        print("Top 50 Search Volume:", top_50_search_volume)
    else:
        print(f"No row found with data-key={target_data_key}")

except Exception as e:
    print("Error extracting data:", str(e))

finally:
    driver.quit()
