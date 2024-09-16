## pip install selenium pandas openpyxl

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import time

# modify this to your local chromedriver path
chrome_driver_path = "C:\\Users\\GAR\\Downloads\\chromedriver-win32\\chromedriver-win32\\chromedriver.exe"

options = webdriver.ChromeOptions()
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)

service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)


url = "https://www.ccilindia.com/web/ccil/security-wise-repo-market-summary"
driver.get(url)

try:
    WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "table.dataTable"))
    )
    print("Table located successfully.")
except Exception as e:
    print("Error locating table:", e)
    driver.quit()

try:
    table = driver.find_element(By.CSS_SELECTOR, "table.dataTable")
    headers = table.find_elements(By.TAG_NAME, "th")
    column_names = [header.text for header in headers]
    print(f"Found {len(column_names)} columns: {column_names}")

    rows = table.find_elements(By.TAG_NAME, "tr")
    data = []

    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")
        if cells:
            row_data = [cell.text for cell in cells]
            data.append(row_data)

    df = pd.DataFrame(data, columns=column_names)
    ## change this depending on where you want your file to show
    df.to_excel("Repo_Market_Summary.xlsx", index=False)
    print("Data has been saved to Repo_Market_Summary.xlsx")

except Exception as e:
    print("Error extracting or saving data:", e)
finally:
    driver.quit()
