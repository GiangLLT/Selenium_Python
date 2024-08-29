import pygetwindow as gw
import pyautogui
import time

def get_active_window_info():
    # Lấy cửa sổ đang hoạt động
    active_window = gw.getActiveWindow()
    
    if active_window is not None:
        print(f"Active Window Title: {active_window.title}")
        print(f"Active Window Size: {active_window.width}x{active_window.height}")
        print(f"Active Window Position: {active_window.left}, {active_window.top}")
    else:
        print("No active window found.")

def perform_actions_on_window():
    # Đảm bảo có thời gian để bạn mở cửa sổ cần thao tác
    time.sleep(5)

    # Lấy cửa sổ đang hoạt động
    active_window = gw.getActiveWindow()

    if active_window is not None:
        # Đặt cửa sổ làm cửa sổ chính
        active_window.activate()
        
        # Chờ một chút để đảm bảo cửa sổ đã được kích hoạt
        time.sleep(1)
        
        # Thực hiện thao tác tự động
        pyautogui.write('Hello, this is an automated action!')
        pyautogui.press('enter')
        
        # Nhấp chuột vào vị trí (x, y) trong cửa sổ
        pyautogui.click(x=active_window.left + 100, y=active_window.top + 100)
    else:
        print("No active window found.")

if __name__ == "__main__":
    get_active_window_info()
    perform_actions_on_window()