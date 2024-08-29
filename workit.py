from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementClickInterceptedException
import pandas as pd
import time
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.edge.options import Options
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os
import shutil


def RPA_workit(data_excel):
    list_data = { 
        'No Data': 0,
        'Tắt kích hoạt Thành Công': 0,
        'Đã Tắt kích hoạt': 0,
        'Kích hoạt thành công': 0,
        'Đã  Kích hoạt': 0,
    }
    data_Nodata = list_data["No Data"]
    data_Unactive = list_data["Tắt kích hoạt Thành Công"]
    data_Unactived = list_data["Đã Tắt kích hoạt"]
    data_Active = list_data["Kích hoạt thành công"]
    data_Actived = list_data["Đã  Kích hoạt"]

    # Chuyển DataFrame thành list các dòng (records)
    records = data_excel.to_dict(orient='records')

    driver = setup_firefox_driver()
    # driver = webdriver.Firefox()

    try:
        # Mở trang web
        driver.get("https://eoffice.bamboocap.com.vn/")

        # Tìm các trường văn bản bằng ID và nhập dữ liệu
        username = driver.find_element(By.ID, "taikhoan")
        password = driver.find_element(By.ID, "matkhau")
        buttonLogin = driver.find_element(By.ID, "btn_Login")

        username.send_keys("INTEGRATION@BAMBOO.COM.VN")
        password.send_keys("Abc@123")
        buttonLogin.click()

        # Đợi cho đến khi trang chuyển hướng sau khi đăng nhập
        WebDriverWait(driver, 10).until(EC.url_contains("LoadContent"))

        print("Đăng nhập web thành công!!!")

        # Mở trang sau khi đăng nhập
        driver.get("https://eoffice.bamboocap.com.vn/LoadContent?rightcode=9502&v=23")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//th[contains(@data-title, "Mã đăng nhập")]/a')))

        print("Bắt đầu xử lý.....")
        for data in records:
            # list_values = list(list_data.values())
            # print(list_values)

            data_stt = str(data["STT"])
            data_email = data["Email"]
            data_status = data["Status"]
            
            time.sleep(1)
            # Tìm và nhấp vào thẻ <a> bên trong <th> có thuộc tính data-title cụ thể
            link = driver.find_element(By.XPATH, '//th[contains(@data-title, "Mã đăng nhập")]/a')
            # Cuộn đến phần tử để đảm bảo nó nằm trong khung nhìn
            driver.execute_script("arguments[0].scrollIntoView();", link)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable(link))

            # Thử nhấp vào phần tử, nếu bị chặn, đợi và thử lại
            try:
                link.click()
            except ElementClickInterceptedException:
                print("Phần tử bị chặn, thử lại...")
                try:
                    button_save_rep = driver.find_element(By.XPATH, '//a[@class="k-button k-button-icontext k-primary k-grid-update"]')
                    button_save_rep.click()
                    time.sleep(1)
                    link.click()
                except Exception as e:
                    time.sleep(1)
                    link.click()

            # Đợi phần tử textbox và button xuất hiện
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "k-textbox")))
            filter_value = driver.find_element(By.CLASS_NAME, "k-textbox")
            filter_button = driver.find_element(By.CLASS_NAME, "k-primary")

            # Cuộn đến phần tử để đảm bảo nó nằm trong khung nhìn
            driver.execute_script("arguments[0].scrollIntoView();", filter_value)

            # Đợi cho đến khi phần tử có thể tương tác được
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.CLASS_NAME, "k-textbox")))

            filter_value.clear()
            filter_value.send_keys(data_email)
            filter_button.click()

            try:
                WebDriverWait(driver, 3).until(
                    EC.presence_of_element_located((By.XPATH, f'//td[text()="{data_email.upper()}"]'))
                )

                # Đợi và nhấp vào nút chỉnh sửa
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'edit')))
                edit_button = driver.find_element(By.NAME, "edit")
                edit_button.click()

                # Đợi phần tử checkbox và các nút lưu/hủy xuất hiện
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "check_active")))
                checkbox_status = driver.find_element(By.ID, "check_active")
                button_save = driver.find_element(By.XPATH, '//a[@class="k-button k-button-icontext k-primary k-grid-update"]')
                button_cancel = driver.find_element(By.XPATH, '//a[@class="k-button k-button-icontext k-grid-cancel"]')

                if data_status == False:
                    if checkbox_status.is_selected():
                        checkbox_status.click()
                        button_save.click()
                        data_Unactive += 1
                        # list_data["Tắt kích hoạt Thành Công"] = data_Unactive
                        print(f"======================================================")
                        print(f"{data_stt} || {data_email} || Tắt kích hoạt Thành Công")
                    else:
                        button_cancel.click()
                        data_Unactived += 1
                        # list_data["Đã Tắt kích hoạt"] = data_Unactived
                        print(f"======================================================")
                        print(f"{data_stt} || {data_email} || Đã Tắt kích hoạt")
                else:
                    if not checkbox_status.is_selected():
                        checkbox_status.click()
                        button_save.click()
                        data_Active += 1
                        # list_data["Kích hoạt thành công"] = data_Active
                        print(f"======================================================")
                        print(f"{data_stt} || {data_email} || Kích hoạt thành công")
                    else:
                        button_cancel.click()
                        data_Actived += 1
                        # list_data["Đã  Kích hoạt"] = data_Actived
                        print(f"======================================================")
                        print(f"{data_stt} || {data_email} || Đã Kích hoạt")
            except Exception as e:
                data_Nodata += 1
                # list_data["No Data"] = data_Nodata
                print(f"======================================================")
                print(f"{data_stt} || {data_email} || No data")
                continue
        
        list_data["Tắt kích hoạt Thành Công"] = data_Unactive
        list_data["Đã Tắt kích hoạt"] = data_Unactived
        list_data["Kích hoạt thành công"] = data_Active
        list_data["Đã  Kích hoạt"] = data_Actived
        list_data["No Data"] = data_Nodata
        print(list_data)
    finally:
        # Đóng trình duyệt
        driver.quit()


def read_excel(file_path, sheet_name=0):
    """
    Đọc dữ liệu từ tệp Excel và trả về DataFrame.

    :param file_path: Đường dẫn tới tệp Excel.
    :param sheet_name: Tên sheet hoặc chỉ số sheet (mặc định là 0 - sheet đầu tiên).
    :return: DataFrame chứa dữ liệu từ tệp Excel.
    """
    try:
        # Đọc dữ liệu từ tệp Excel
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        return df
    except Exception as e:
        print(f"Đã xảy ra lỗi khi đọc tệp Excel: {e}")
        return None

def setup_chrome_driver():
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')

    driver = webdriver.Chrome(options=chrome_options)
    return driver

def setup_firefox_driver():
    firefox_options = webdriver.FirefoxOptions()
    firefox_options.add_argument('--headless')
    firefox_options.add_argument('--disable-gpu')
    firefox_options.add_argument('--no-sandbox')
    firefox_options.add_argument('--disable-dev-shm-usage')

    driver = webdriver.Firefox(options=firefox_options)
    return driver

def setup_edge_driver():
    edge_options =webdriver.EdgeOptions()
    edge_options.add_argument('--headless')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')

    driver = webdriver.Edge(options=edge_options)
    return driver


if __name__ == "__main__":
    # Yêu cầu người dùng nhập đường dẫn tới tệp Excel
    file_path = input("Vui lòng nhập đường dẫn tới tệp Excel: ")

    # Đọc dữ liệu từ tệp Excel
    try:
        df = read_excel(file_path)

        # Hiển thị dữ liệu đọc được
        if df is not None:
            RPA_workit(df)
        else:
            print("No Data")
    except ValueError:
        print("Read Excel Error")

    input("Nhấn phím bất kỳ để kết thúc...")
