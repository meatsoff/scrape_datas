import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import csv

# Khởi tạo trình duyệt
driver = webdriver.Chrome()
driver.maximize_window()  # Mở rộng cửa sổ trình duyệt
csv_file = open('Output_scraped.csv', 'w', newline='', encoding='utf-8-sig')
csv_writer = csv.writer(csv_file)

try:
    # Mở trang web
    driver.get('https://dichvucong.dav.gov.vn/congbothuoc/index')
    source = driver.page_source
    soup = BeautifulSoup(source, 'html.parser')

    # Chờ trang tải xong (sử dụng WebDriverWait)
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder*="tìm kiếm"]'))
    )

    # Đến ô tìm kiếm và điền thông tin rồi Enter (có thể lấy từ file excel)
    input_field = driver.find_element(By.CSS_SELECTOR, 'input[placeholder*="tìm kiếm"]')
    input_field.send_keys('Paracetamol')
    input_field.send_keys(Keys.RETURN)

    time.sleep(2)
    # Lấy ra số kết quả tìm thấy
    result_count_element = driver.find_element(By.CSS_SELECTOR, 'strong.ng-binding')
    result_count = result_count_element.text.strip()  # Clean any extra whitespace

    # Click hiển thị 20 cái/ lần
    span_element = driver.find_element(By.CSS_SELECTOR, 'span.k-input.ng-scope')
    span_element.click()
    option = driver.find_element(By.XPATH, '//li[@data-offset-index="3"]')
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable(option)
    )
    option.click()
    time.sleep(2)

    # Tính số lần chuyển trang
    result_count = result_count.replace(",", "")
    loop_times = int(result_count) / 20 + 1

    # Lấy và ghi vào excel 20 cái mỗi lần chuyển trang
    for i in range(int(loop_times)):
        grid_content = driver.find_element(By.CSS_SELECTOR, 'div.k-grid-content.k-auto-scrollable')
        for content_each_row in grid_content.find_elements(By.TAG_NAME, 'tr'):
            data_cells = content_each_row.find_elements(By.TAG_NAME, 'td')
            detail_data_col = [data_cell.text.strip() for data_cell in data_cells]
            print(detail_data_col)
            csv_writer.writerows([detail_data_col])
        if i != int(loop_times):
            span_element = driver.find_element(By.CSS_SELECTOR, 'span.k-icon.k-i-arrow-60-right')
            span_element.click()
            time.sleep(2)

    csv_file.close()


except TimeoutException:
    print("Timeout khi chờ phần tử xuất hiện")
except Exception as e:
    print(f"Lỗi: {str(e)}")
finally:
    driver.quit()