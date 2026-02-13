import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import pandas as pd

df = pd.read_csv("NovoPro_input.csv")
id_smiles = df["ID"].dropna().tolist()
results = []

print(id_smiles)

driver = webdriver.Chrome()
driver.maximize_window()

try:
    # Mở trang web
    driver.get('https://www.novoprolabs.com/tools/convert-peptide-to-smiles-string')
    source = driver.page_source
    soup = BeautifulSoup(source, 'html.parser')

    # Chờ trang tải xong (sử dụng WebDriverWait)
    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="input-sequence"]'))
    )

    for sid in id_smiles:
        # Tới input để clear nội dung và điền ID mới vào
        input_field = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="input-sequence"]'))
        )
        input_field.clear()
        input_field.send_keys(sid)
        # Submit
        driver.find_element(By.XPATH, '//*[@id="input-design-submit"]').send_keys(Keys.RETURN)
        # Chờ và lấy kết quả
        WebDriverWait(driver, 20).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, '//*[@id="output-res"]/textarea'))
            )
        )
        result_search = driver.find_element(By.XPATH,'//*[@id="output-res"]/textarea').text
        print(result_search)
        # Ghi vào excel
        results.append({"ID": sid, "Code SMILES": result_search})
        time.sleep(3)

except TimeoutException:
    print("Timeout error!")
except Exception as e:
    print(f"Lỗi: {str(e)}")
finally:
    driver.quit()

pd.DataFrame(results).to_excel("NovoPro_output.xlsx", index=False)
print("Finished saving datas in SEA_output.xlsx")