import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd

df = pd.read_excel("NovoPro_input_5k.xlsx")
id_smiles = df["ID"].dropna().tolist()
results = []

print(f"Tổng số ID: {len(id_smiles)}")

driver = webdriver.Chrome()
driver.maximize_window()

try:
    driver.get('https://www.novoprolabs.com/tools/convert-peptide-to-smiles-string')

    wait = WebDriverWait(driver, 20)

    # Chờ input load
    input_field = wait.until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="input-sequence"]'))
    )

    for numid, sid in enumerate(id_smiles, start=1):
        print(f"Đang xử lý {numid}/{len(id_smiles)} - {sid}")

        input_field.clear()
        input_field.send_keys(sid)

        # Chờ overlay biến mất
        try:
            wait.until(
                EC.invisibility_of_element_located((By.CLASS_NAME, "layui-layer-shade"))
            )
        except:
            pass

        submit_btn = wait.until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="input-design-submit"]'))
        )
        submit_btn.click()

        output_area = wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="output-res"]/textarea'))
        )

        result_search = output_area.get_attribute("value")

        print(f"Kết quả: {result_search}")

        results.append({
            "ID": sid,
            "Code SMILES": result_search
        })

except TimeoutException:
    print("Timeout error!")
except Exception as e:
    print(f"Lỗi: {str(e)}")
finally:
    driver.quit()

pd.DataFrame(results).to_excel("NovoPro_output.xlsx", index=False)
print("Finished!")
