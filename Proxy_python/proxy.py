import requests
from bs4 import BeautifulSoup
import os
import subprocess

def list_proxy(url):
    # URL của trang web cung cấp danh sách proxy miễn phí
    # url = "https://free-proxy-list.net/"

    # Gửi yêu cầu HTTP GET để lấy nội dung trang web
    response = requests.get(url)
    html = response.content

    # Phân tích nội dung HTML bằng BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.find('table', {'class': 'table table-striped table-bordered'})
    rows = table.tbody.find_all('tr')

    # Duyệt qua các hàng trong bảng và lấy thông tin proxy
    proxies = []
    for row in rows[:18]:
        cols = row.find_all('td')
        ip = cols[0].text
        port = cols[1].text
        code = cols[2].text
        country = cols[3].text
        anonymity = cols[4].text
        https = cols[6].text
        proxy = {
            'ip': ip,
            'port': port,
            'code': code,
            'country': country,
            'anonymity': anonymity,
            'https': https
        }
        proxies.append(proxy)
    return proxies

# Hàm kiểm tra proxy
def check_proxy(proxy):
    try:
        proxy_url = f"http://{proxy['ip']}:{proxy['port']}"
        proxies = {
            "http": proxy_url,
            "https": proxy_url
        }
        response = requests.get("http://www.google.com", proxies=proxies, timeout=5)
        print(f"{proxy['ip']} is working")
        return response.status_code == 200
    except:
        print(f"{proxy['ip']} is not working")
        return False

# Hàm kiểm tra proxy với nhiều URL
def check_proxy_multiple(proxy):
    test_urls = ["http://www.google.com", "http://www.example.com"]
    proxy_url = f"http://{proxy['ip']}:{proxy['port']}"
    proxies = {
        "http": proxy_url,
        "https": proxy_url
    }
    for url in test_urls:
        try:
            response = requests.get(url, proxies=proxies, timeout=5)
            if response.status_code == 200:
                return True
        except:
            continue
    return False

def connect_via_proxy1():
    # Đặt địa chỉ proxy và port
    proxy_address = "216.137.184.253"
    proxy_port = "80"

    # Đường dẫn đầy đủ đến PowerShell
    powershell_path = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"

    # Tạo chuỗi lệnh để thay đổi thiết lập proxy
    set_proxy_command = f'Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings" -Name ProxyServer -Value "{proxy_address}:{proxy_port}"'
    enable_proxy_command = 'Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings" -Name ProxyEnable -Value 1'

    # Thực thi các lệnh thay đổi thiết lập proxy bằng PowerShell
    data = f'{powershell_path} -Command "{set_proxy_command}"'
    data1 = f'{powershell_path} -Command "{enable_proxy_command}"'
    print(data)
    print(data1)
    os.system(f'{powershell_path} -Command "{set_proxy_command}"')
    os.system(f'{powershell_path} -Command "{enable_proxy_command}"')

    print("Proxy settings have been updated.")

def connect_via_proxy():
    # Đặt địa chỉ proxy và port
    proxy_address = "216.137.184.253"
    proxy_port = "80"

    # Đường dẫn đầy đủ tới PowerShell
    powershell_path = r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'

    # Tạo chuỗi lệnh để thay đổi thiết lập proxy
    set_proxy_command = f'Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings" -Name ProxyServer -Value "{proxy_address}:{proxy_port}"'
    enable_proxy_command = 'Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings" -Name ProxyEnable -Value 1'

    # Thực thi các lệnh thay đổi thiết lập proxy bằng PowerShell
    subprocess.run([powershell_path, '-Command', set_proxy_command])
    subprocess.run([powershell_path, '-Command', enable_proxy_command])

    print("Proxy settings have been updated.")

def disconnect_proxy():
    # Đường dẫn đầy đủ tới PowerShell
    powershell_path = r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'

    # Tạo chuỗi lệnh để tắt proxy
    disable_proxy_command = 'Set-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings" -Name ProxyEnable -Value 0'
    clear_proxy_command = 'Remove-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings" -Name ProxyServer'

    # Thực thi các lệnh để tắt proxy bằng PowerShell
    subprocess.run([powershell_path, '-Command', disable_proxy_command])
    subprocess.run([powershell_path, '-Command', clear_proxy_command])

    print("Proxy settings have been reset to normal.")

if __name__ == "__main__":
    try:
        connect_via_proxy()
        # Kiểm tra danh sách proxy và lưu các proxy hoạt động vào danh sách mới
        # working_proxies = []
        # url = "https://free-proxy-list.net/"
        # proxies = list_proxy(url)
        # if proxies:
        #     for proxy in proxies:
        #         if check_proxy(proxy):
        #             working_proxies.append(proxy)

        #     # In danh sách proxy hoạt động
        #     print("Working proxies:")
        #     for proxy in working_proxies:
        #         print(proxy)
        # else:
        #     print("Proxy is Empty")
    except ValueError as e:
        print(f"Error: {e}")

    # Ngắt kết nối proxy
    disconnect_proxy()

    input("Press any key to exit...")
