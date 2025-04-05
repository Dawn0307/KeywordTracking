import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 將欄位字母（如 "AC"）轉成數字
def col_letter_to_number(col_letter):
    col_letter = col_letter.upper()
    result = 0
    for char in col_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result

# 初始化 Google Sheets API
def initialize_google_sheets():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(credentials)
    return client

# ✅ 批量寫入公式（避免超過 API 限制）
def batch_update_formulas(sheet, target_row, formula_sheet):
    requests = []
    for col_letter, formula in formula_sheet.items():
        try:
            formula_with_row = formula.replace("{num}", str(target_row))
            cell_label = f"{col_letter}{target_row}"
            # ✅ 不再加 "="，因為你的公式本身已經有開頭的 "="
            requests.append({
                "range": cell_label,
                "values": [[formula_with_row]]
            })
        except Exception as e:
            print(f"❌ Failed to prepare {col_letter}{target_row}: {e}")

    if requests:
        body = {
            "valueInputOption": "USER_ENTERED",
            "data": requests
        }
        sheet.spreadsheet.values_batch_update(body)
        print(f"✅ Formulas batch-updated for row {target_row}")


# 📄 公式表
formula_sheet = {
    "BA": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!Y:Y) /100",
    "BB": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!T:T)",
    "BC": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!V:V)",
    "BD": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!U:U)",
    "BE": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AA:AA) /100",
    "CG": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AJ:AJ) /100",
    "CH": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AE:AE)",
    "CI": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AG:AG)",
    "CJ": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AF:AF)",
    "CK": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AL:AL) /100",
    "BI": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AU:AU) /100",
    "BJ": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AP:AP)",
    "BK": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AR:AR)",
    "BL": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AQ:AQ)",
    "BM": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!AW:AW) /100",
    "CW": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!BF:BF) /100",
    "CX": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!BA:BA)",
    "CY": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!BC:BC)",
    "CZ": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!BB:BB)",
    "DA": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!BH:BH) /100",
    "DM": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!BQ:BQ) /100",
    "DN": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!BL:BL)",
    "DO": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!BN:BN)",
    "DP": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!BM:BM)",
    "DQ": "=SUMIF(ADS_CLARITY!$S:$S,$D{num},ADS_CLARITY!BO:BO) /100"
}

# 更新 Google Sheets
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
        batch_update_formulas(sheet, target_row, formula_sheet)
    else:
        print(f"⚠️ Date {date} not found in the sheet.")

# 啟動 Chrome
def start_chrome_with_existing_profile():
    options = Options()
    user_data_dir = 'C:\\Users\\A\\AppData\\Local\\Google\\Chrome\\User Data'
    profile_dir = 'Default'
    options.add_argument(f"user-data-dir={user_data_dir}")
    options.add_argument(f"profile-directory={profile_dir}")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

# 主爬蟲邏輯
def scrape_data():
    driver = start_chrome_with_existing_profile()
    url = 'https://members.helium10.com/keyword-tracker?accountId=1545862407&per-page=50'
    driver.get(url)

    target_data_key_set = ['1615430','1126568','1771901','1556573','1884995','1865893','1737627','1825151','2117935','1874489','1818149','1870601','1909877']
    target_data_key_col = [2, 10, 18, 26, 34, 50, 58, 74, 82, 90, 98, 106, 114]

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
                update_google_sheet("ClarityCreation-OPS-2025", current_date, target_data_key, col, top_10, top_50)
            else:
                print(f"⚠️ No row found for key: {target_data_key}")

    except Exception as e:
        print("❌ Error extracting data:", str(e))

    finally:
        driver.quit()

# 執行主程式
if __name__ == "__main__":
    scrape_data()
