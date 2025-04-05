import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
from datetime import datetime, timedelta
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 1. 初始化 Google Sheets API
def initialize_google_sheets():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(credentials)
    return client


def col_letter_to_number(col):
    num = 0
    for c in col:
        num = num * 26 + (ord(c.upper()) - ord('A') + 1)
    return num
# 2. 更新 Google Sheets
def update_google_sheet(sheet_name, date, product_1_data, product_2_data,product_3_data):
    client = initialize_google_sheets()
    sheet = client.open(sheet_name).worksheet("DailyUpdates")

    # 找到目標日期對應的行
    date_col = 4  # 日期在第 D 欄 (第 4 欄)
    rows = sheet.get_all_values()
    target_row = None

    for idx, row in enumerate(rows):
        if len(row) >= date_col and row[date_col - 1] == date:
            target_row = idx + 1  # gspread 的行索引從 1 開始
            break

    # 定義公式（保持不變）
    formula_sheet = {

        "J": "=(SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!AF:AF) + SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!BM:BM))/200",
        "K": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!AH:AH) + SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!BH:BH)",
        "L": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!BJ:BJ)",
        "M": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!AI:AI) + SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!BI:BI)",
        "N": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!BO:BO) /100",

        "R": "=(SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!BW:BW) + SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!DD:DD))/200",
        "S": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!BY:BY) + SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!CY:CY)",
        "T": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!DA:DA)",
        "U": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!BZ:BZ) + SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!CZ:CZ)",
        "V": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!DF:DF) /100",


        #Travelbag 只有一個，沒有SB
        "AH": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!DO:DO) /100",
        "AI": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!DJ:DJ)",
        "AJ": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!DL:DL)",
        "AK": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!DK:DK)",
        "AL": "=SUMIF(ADS_Vortex!$AA:$AA,$D{num},ADS_Vortex!DQ:DQ) /100",

    }

    if target_row:
        # **更新第一個商品的數據** (B, C 欄)
        sheet.update_cell(target_row, 2, product_1_data[0])  # Top 10 (B欄)
        sheet.update_cell(target_row, 3, product_1_data[1])  # Top 50 (C欄)

        # **更新第二個商品的數據** (O, P 欄)
        sheet.update_cell(target_row, 15, product_2_data[0])  # Top 10 (O欄)
        sheet.update_cell(target_row, 16, product_2_data[1])  # Top 50 (P欄)

        # **更新第4  個商品的數據** (O, P 欄)
        sheet.update_cell(target_row, 31, product_3_data[0])  # Top 10 (O欄)
        sheet.update_cell(target_row, 32, product_3_data[1])  # Top 50 (P欄)

        print(f"Data updated successfully for date: {date}")

        # **寫入原有的公式**
        for col_letter, formula in formula_sheet.items():
            formula_with_row = formula.replace("{num}", str(target_row))
            col_number = col_letter_to_number(col_letter)
            sheet.update_cell(target_row, col_number, formula_with_row)
            print(f"Formula written to {col_letter}{target_row}: {formula_with_row}")

    else:
        print(f"Date {date} not found in the sheet.")


# 3. 爬取數據邏輯
def start_chrome_with_existing_profile():
    options = Options()
    user_data_dir = 'C:\\Users\\A\\AppData\\Local\\Google\\Chrome\\User Data'
    profile_dir = 'Default'

    options.add_argument(f"user-data-dir={user_data_dir}")
    options.add_argument(f"profile-directory={profile_dir}")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    
    return driver

def scrape_data():
    driver = start_chrome_with_existing_profile()
    url = 'https://members.helium10.com/keyword-tracker?accountId=1545862407'
    driver.get(url)

    try:
        # 等待表格加載
        rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tr.kt-orders-row"))
        )
        
        # 目標 data-key
        target_data_keys = ["1988788", "2131278","2138431"]
        extracted_data = {}

        for target_data_key in target_data_keys:
            target_row = None
            for row in rows:
                if row.get_attribute("data-key") == target_data_key:
                    target_row = row
                    break

            if target_row:
                top_10_search_volume = target_row.find_element(By.XPATH, ".//div[@class='col-md-6 search_volume_keywords']//p[text()='Top 10']/preceding-sibling::span[1]").text
                top_50_search_volume = target_row.find_element(By.XPATH, ".//div[@class='col-md-6 search_volume_keywords']//p[text()='Top 50']/preceding-sibling::span[1]").text
                
                extracted_data[target_data_key] = (top_10_search_volume, top_50_search_volume)
            else:
                print(f"No row found with data-key={target_data_key}")

        # 提取 **前一天** 的日期
        current_date = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

        # 更新 Google Sheets
        update_google_sheet(
            "BV_OPS_2025",
            current_date,
            extracted_data.get("1988788", (None, None)),  # 第一個商品的數據
            extracted_data.get("2131278", (None, None)),   # 第二個商品的數據
            extracted_data.get("2138431", (None, None))
        )

    except Exception as e:
        print("Error extracting data:", str(e))

    finally:
        driver.quit()

# 4. 執行主程式
if __name__ == "__main__":
    scrape_data()
