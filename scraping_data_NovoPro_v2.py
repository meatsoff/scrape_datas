import os
import time
import random
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
    ElementClickInterceptedException,
    WebDriverException
)

# ==========================
# CONFIG
# ==========================
INPUT_FILE = "NovoPro_input_5k.xlsx"
OUTPUT_FILE = "NovoPro_output.xlsx"
AUTO_SAVE_EVERY = 50
RESTART_EVERY = 300
MAX_RETRY = 3
URL = "https://www.novoprolabs.com/tools/convert-peptide-to-smiles-string"

# ==========================
# LOAD INPUT
# ==========================
df = pd.read_excel(INPUT_FILE)
id_smiles = df["ID"].dropna().tolist()

# ==========================
# RESUME LOGIC
# ==========================
if os.path.exists(OUTPUT_FILE):
    old_df = pd.read_excel(OUTPUT_FILE)
    done_ids = set(old_df["ID"].tolist())
    results = old_df.to_dict("records")
    print(f"🔁 Resume: đã xử lý {len(done_ids)} dòng")
else:
    done_ids = set()
    results = []
    print("🆕 Chạy mới hoàn toàn")

# ==========================
# DRIVER FACTORY
# ==========================
def create_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = webdriver.Chrome(options=options)
    driver.get(URL)
    return driver

def is_driver_alive(driver):
    try:
        driver.current_url
        return True
    except:
        return False

# ==========================
# START
# ==========================
driver = create_driver()
wait = WebDriverWait(driver, 30)

try:
    wait.until(EC.presence_of_element_located((By.ID, "input-sequence")))

    for numid, sid in enumerate(id_smiles, start=1):

        if sid in done_ids:
            continue

        print(f"▶ {numid}/{len(id_smiles)} - {sid}")

        # Restart định kỳ
        if numid % RESTART_EVERY == 0:
            print("🔄 Restart driver định kỳ")
            driver.quit()
            driver = create_driver()
            wait = WebDriverWait(driver, 30)

        # Nếu driver chết thì tạo lại
        if not is_driver_alive(driver):
            print("⚠ Driver chết → khởi động lại")
            driver = create_driver()
            wait = WebDriverWait(driver, 30)

        success = False

        for attempt in range(MAX_RETRY):
            try:
                input_field = wait.until(
                    EC.presence_of_element_located((By.ID, "input-sequence"))
                )
                input_field.clear()
                input_field.send_keys(sid)

                # Chờ overlay biến mất
                try:
                    wait.until(
                        EC.invisibility_of_element_located(
                            (By.CLASS_NAME, "layui-layer-shade")
                        )
                    )
                except:
                    pass

                submit_btn = wait.until(
                    EC.element_to_be_clickable((By.ID, "input-design-submit"))
                )

                driver.execute_script(
                    "arguments[0].scrollIntoView(true);", submit_btn
                )
                driver.execute_script(
                    "arguments[0].click();", submit_btn
                )

                output_area = wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, '//*[@id="output-res"]/textarea')
                    )
                )

                result_search = output_area.get_attribute("value")

                results.append({
                    "ID": sid,
                    "Code SMILES": result_search
                })

                done_ids.add(sid)
                success = True
                print("✔ Thành công")
                break

            except (
                TimeoutException,
                StaleElementReferenceException,
                ElementClickInterceptedException,
                WebDriverException
            ):
                print(f"⚠ Retry {attempt+1}/{MAX_RETRY}")
                time.sleep(2)

        if not success:
            print("❌ Lỗi → ghi ERROR")
            results.append({
                "ID": sid,
                "Code SMILES": "ERROR"
            })

        # Auto save
        if numid % AUTO_SAVE_EVERY == 0:
            pd.DataFrame(results).to_excel(OUTPUT_FILE, index=False)
            print("💾 Auto saved")

        time.sleep(random.uniform(1, 2))

except Exception as e:
    print(f"🔥 Lỗi hệ thống: {str(e)}")

finally:
    driver.quit()

pd.DataFrame(results).to_excel(OUTPUT_FILE, index=False)
print("🎉 HOÀN TẤT!")