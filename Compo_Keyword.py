import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime

# 將欄位字母（如 "AC"）轉換成對應欄位數字
def col_letter_to_number(col_letter):
    col_letter = col_letter.upper()
    result = 0
    for char in col_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result

# 1. 初始化 Google Sheets API
def initialize_google_sheets():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(credentials)
    return client

# 公式設定
formula_sheet = {

    "E": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!Y:Y) /100",
    "F": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!T:T)",
    "G": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!V:V)",
    "H": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!U:U)",
    "I": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AA:AA) /100",

    "M": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AJ:AJ) /100",
    "N": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AE:AE)",
    "O": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AG:AG)",
    "P": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AF:AF)",
    "Q": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AL:AL) /100",

    "U": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AU:AU) /100",
    "V": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AP:AP)",
    "W": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AR:AR)",
    "X": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AQ:AQ)",
    "Y": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!AW:AW) /100",

    "AC": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!BF:BF) /100",
    "AD": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!BA:BA)",
    "AE": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!BC:BC)",
    "AF": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!BB:BB)",
    "AG": "=SUMIF(ADS_Componova!$S:$S,$D{num},ADS_Componova!BH:BH) /100"
}

# 2. 更新 Google Sheets 資料
def update_google_sheet(sheet_name, date, target_data_key, col, top_10_volume, top_50_volume):
    client = initialize_google_sheets()
    sheet = client.open(sheet_name).worksheet("DailyUpdates")

    date_col = 4
    rows = sheet.get_all_values()
    target_row = None

    for idx, row in enumerate(rows):
        if len(row) >= date_col and row[date_col - 1] == date:
            target_row = idx + 1
            break

    if target_row:
        sheet.update_cell(target_row, col, top_10_volume)
        sheet.update_cell(target_row, col + 1, top_50_volume)
        print(f"✅ Data updated for key {target_data_key} on date: {date}")

        # 寫入原有的公式
        for col_letter, formula in formula_sheet.items():
            try:
                formula_with_row = formula.replace("{num}", str(target_row))
                col_number = col_letter_to_number(col_letter)
                sheet.update_cell(target_row, col_number, formula_with_row)
                print(f"Formula written to {col_letter}{target_row}: {formula_with_row}")
            except Exception as e:
                print(f"❌ Failed to update {col_letter}{target_row}: {e}")
    else:
        print(f"⚠️ Date {date} not found in the sheet.")

# 3. 啟動 Chrome 瀏覽器（使用本機個人資料）
def start_chrome_with_existing_profile():
    options = Options()
    user_data_dir = 'C:\\Users\\A\\AppData\\Local\\Google\\Chrome\\User Data'
    profile_dir = 'Default'
    options.add_argument(f"user-data-dir={user_data_dir}")
    options.add_argument(f"profile-directory={profile_dir}")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

# 4. 主爬蟲流程
def scrape_data():
    driver = start_chrome_with_existing_profile()
    url = 'https://members.helium10.com/keyword-tracker?accountId=1545862407&per-page=50'
    driver.get(url)

    target_data_key_set = ['2099291', '2085904', '2085905', '2088793']
    target_data_key_col = [2, 10, 18, 26]

    try:
        rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tr.kt-orders-row"))
        )

        current_date = datetime.now().strftime('%Y-%m-%d')

        for target_data_key, col in zip(target_data_key_set, target_data_key_col):
            target_row = next((r for r in rows if r.get_attribute("data-key") == target_data_key), None)

            if target_row:
                top_10 = target_row.find_element(By.XPATH, ".//div[@class='col-md-6 search_volume_keywords']//p[text()='Top 10']/preceding-sibling::span[1]").text
                top_50 = target_row.find_element(By.XPATH, ".//div[@class='col-md-6 search_volume_keywords']//p[text()='Top 50']/preceding-sibling::span[1]").text

                print(f"🔍 Key {target_data_key} - Top 10: {top_10}, Top 50: {top_50}")
                update_google_sheet("CompoNova-OPS-2025", current_date, target_data_key, col, top_10, top_50)
            else:
                print(f"⚠️ No row found for key: {target_data_key}")

    except Exception as e:
        print("❌ Error extracting data:", str(e))

    finally:
        driver.quit()

# 5. 執行主程式
if __name__ == "__main__":
    scrape_data()
