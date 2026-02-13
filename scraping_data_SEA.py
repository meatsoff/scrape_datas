import os
import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import pandas as pd
# tải bằng requests
import requests

df = pd.read_excel("SEA_input.xlsx")
code_smiles = df["Code SMILES"].dropna().tolist()
id_smiles = df["ID"].dropna().tolist()
results = []

print(code_smiles)

driver = webdriver.Chrome()
driver.maximize_window()

try:
    # Mở trang web
    driver.get('https://sea.bkslab.org/')
    source = driver.page_source
    soup = BeautifulSoup(source, 'html.parser')

    # Chờ trang tải xong (sử dụng WebDriverWait)
    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder*="Paste SMILES or try the example below"]'))
    )

    for code, sid in zip(code_smiles, id_smiles):
        input_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder*="Paste SMILES"]'))
        )
        input_field.clear()
        input_field.send_keys(code)
        input_field.send_keys(Keys.RETURN)

        WebDriverWait(driver, 20).until(
            EC.any_of(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'table.table-bordered')),
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.alert.alert-error'))
            )
        )

        if driver.find_elements(By.CSS_SELECTOR, 'table.table-bordered'):
            result_search = driver.find_element(By.XPATH,
                            "//h1[contains(@class, 'clearfix')]//a[2]").text
            results.append({"Code SMILES": code, "ID": sid, "Download files": result_search})
            time.sleep(1)
            # Tìm và tải file
            download_btn = driver.find_element(By.CSS_SELECTOR, "a[href$='.zip']")
            file_url = download_btn.get_attribute("href")
            print("Download URL:", file_url)
            filename = os.path.basename(file_url)
            file_path = os.path.join(r'C:\Users\Hi\Downloads', filename)
            response = requests.get(file_url, timeout=30)
            if response.status_code == 200:
                with open(file_path, "wb") as f:
                    f.write(response.content)
            print("Downloaded:", filename)
            time.sleep(5)

        elif driver.find_elements(By.XPATH, '/html/body/div/div/div/div'):
            results.append({"Code SMILES": code, "Download files": 'Not found to download'})
        driver.get('https://sea.bkslab.org/')

except TimeoutException:
    print("Timeout error!")
except Exception as e:
    print(f"Lỗi: {str(e)}")
finally:
    driver.quit()

pd.DataFrame(results).to_excel("SEA_output.xlsx", index=False)
print("Finished saving datas in SEA_output.xlsx")