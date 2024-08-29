import os
import subprocess
import logging
import time
import keyboard
from pywinauto import Desktop
from pywinauto.application import Application
from pywinauto import mouse
from pywinauto.keyboard import send_keys
import psutil
import pygetwindow as gw
from pynput import mouse, keyboard as kb
import json
import win32gui
import win32process
import pefile
import pyautogui
import re
import winreg as reg

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
# os.environ['PYDEVD_WARN_SLOW_RESOLVE_TIMEOUT'] = '10.0'

def is_registry_installed():
    try:
        registry_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        registry = reg.ConnectRegistry(None, reg.HKEY_LOCAL_MACHINE)
        uninstall_key = reg.OpenKey(registry, registry_path)
        return uninstall_key
    except Exception as e:
        print(f"Error checking registry: {e}")
    return False

def is_ultraviewer_installed(uninstall_key,product_name):
    # Kiểm tra khóa registry để xem UltraViewer có được cài đặt không
    try:
        # registry_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        # registry = reg.ConnectRegistry(None, reg.HKEY_LOCAL_MACHINE)
        # uninstall_key = reg.OpenKey(registry, registry_path)
        
        for i in range(0, reg.QueryInfoKey(uninstall_key)[0]):
            sub_key_name = reg.EnumKey(uninstall_key, i)
            sub_key = reg.OpenKey(uninstall_key, sub_key_name)
            try:
                display_name, _ = reg.QueryValueEx(sub_key, "DisplayName")
                if product_name.lower() in display_name.lower():
                    return True
            except FileNotFoundError:
                continue
    except Exception as e:
        print(f"Error checking registry: {e}")
    return False

def get_installed_software():
    global installed_software_cache
    if installed_software_cache is not None:
        return installed_software_cache

    logging.info("Loading installed software data...")
    try:
        output = subprocess.check_output(['wmic', 'product', 'get', 'name'], universal_newlines=True)
        installed_software_cache = output.lower().split('\n')
        logging.info("Installed software data loaded successfully.")
    except subprocess.CalledProcessError:
        installed_software_cache = []
        logging.error("Error loading installed software data.")
    print(installed_software_cache)
    return installed_software_cache

def is_software_installed(software_name, installed_cache,product_name):

    for software in installed_cache:
        # if software_name.lower() in software:
        if product_name.lower() in software:
            return True
    return False


def run_script(script_file):
    with open(script_file, 'r') as f:
        actions = f.readlines()
    for action in actions:
        exec(action.strip())

def enum_windows_callback(hwnd, windows):
    """Hàm callback để thêm cửa sổ và PID vào danh sách."""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        window_title = win32gui.GetWindowText(hwnd)
        windows.append((hwnd, pid, window_title))
    except Exception as e:
        print(f"Error getting PID for window: {e}")

def get_windows_by_pid(target_pid, product_name):
    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    # return [(hwnd, window_title) for hwnd, pid, window_title in windows if product_name.lower() in window_title.lower() and pid == target_pid]
    return [(hwnd, window_title) for hwnd, pid, window_title in windows if product_name.lower() in window_title.lower()]

def get_window_by_title(title):
    def callback(hwnd, windows):
        if win32gui.IsWindowVisible(hwnd) and title in win32gui.GetWindowText(hwnd):
            windows.append(hwnd)
    windows = []
    win32gui.EnumWindows(callback, windows)
    return windows

def get_window_by_id(product_name, use_pid=False, pid=None):
    # Kiểm tra xem có sử dụng PID để lấy cửa sổ hay không
    if use_pid and pid:
        window = None
        # print(pid)
        windows = get_windows_by_pid(pid, product_name)
        # windows1 = gw.getActiveWindow()
        # print("active window:" + str(windows1._hWnd))
        time.sleep(1.0)
        if windows:
            for hwnd,title  in windows:
                if win32gui.IsWindowVisible(hwnd):
                    # window = gw.getWindowsWithTitle(str(hwnd[1]))[0]
                    window = gw.getWindowsWithTitle(title)[0]
                    rect = win32gui.GetWindowRect(hwnd)
                    logging.info(f"Window: {title}, HWND: {hwnd}, Size: {rect[2]-rect[0]}x{rect[3]-rect[1]}")
                    # print("window:" + str(window._hWnd))
                    break
        else:
            print("No window found for the given PID.")
    else:
        app = gw.getActiveWindow()
        window = gw.getWindowsWithTitle(str(app.title))[0]

    return window

# def record_user_actions(product_name, use_pid=False, pid=None):
def record_user_actions1(window,script_file):
    print("Recording user actions. Press 'Esc' to stop recording.")
    actions = []
    start_time = time.time()
    stop_recording = [False]

    def stop():
        stop_recording[0] = True

    keyboard.add_hotkey('q', stop)

    if window is None:
        print("No active window found.")
        return

    print(f"Recording actions for window: {window.title}")

    def on_click(x, y, button, pressed):
        if window.isActive:
            relative_x = x - window.left
            relative_y = y - window.top
            action = {
                "type": "click",
                "x": relative_x,
                "y": relative_y,
                "time": time.time() - start_time,
                "button": str(button),
                "pressed": pressed
            }
            actions.append(action)
            print(f"Recorded click at ({relative_x}, {relative_y}) with button {button}, pressed: {pressed}")

    def on_key_press(key):
        if window.isActive:
            try:
                key_str = key.char
            except AttributeError:
                key_str = str(key)
            action = {
                "type": "keypress",
                "keys": key_str,
                "time": time.time() - start_time
            }
            actions.append(action)
            print(f"Recorded keypress: {key_str}")

    def on_key_release(key):
        if key == kb.Key.esc:
            stop()

    mouse_listener = mouse.Listener(on_click=on_click)
    keyboard_listener = kb.Listener(on_press=on_key_press, on_release=on_key_release)

    mouse_listener.start()
    keyboard_listener.start()

    while not stop_recording[0]:
        time.sleep(0.1)

    mouse_listener.stop()
    keyboard_listener.stop()

    keyboard.remove_hotkey('q')

    directory = os.path.dirname(script_file)
    if not os.path.exists(directory):
        os.makedirs(directory)

    # Lưu các hành động vào file JSON
    with open(script_file, 'w') as f:
        for action in actions:
            f.write(action + '\n')

    print(f"Recorded user actions saved to {script_file}")

def record_user_actions1(window, script_file):
    print("Recording user actions. Press 'Esc' to stop recording.")
    actions = []
    start_time = time.time()
    stop_recording = [False]

    def stop():
        stop_recording[0] = True

    kb.Listener.suppress_hotkeys = True

    if window is None:
        print("No active window found.")
        return

    print(f"Recording actions for window: {window.title}")

    def on_click(x, y, button, pressed):
        if window.isActive:
            relative_x = x - window.left
            relative_y = y - window.top
            action = {
                "type": "click",
                "x": relative_x,
                "y": relative_y,
                "time": time.time() - start_time,
                "button": str(button),
                "pressed": pressed
            }
            actions.append(action)
            print(f"Recorded click at ({relative_x}, {relative_y}) with button {button}, pressed: {pressed}")

    def on_key_press(key):
        if window.isActive:
            try:
                key_str = key.char
            except AttributeError:
                key_str = str(key)
            action = {
                "type": "keypress",
                "keys": key_str,
                "time": time.time() - start_time
            }
            actions.append(action)
            print(f"Recorded keypress: {key_str}")

    def on_key_release(key):
        if key == kb.Key.esc:
            stop_recording[0] = True
            return False

    mouse_listener = mouse.Listener(on_click=on_click)
    keyboard_listener = kb.Listener(on_press=on_key_press, on_release=on_key_release)

    mouse_listener.start()
    keyboard_listener.start()

    while not stop_recording[0]:
        time.sleep(0.1)

    mouse_listener.stop()
    keyboard_listener.stop()

    directory = os.path.dirname(script_file)
    if not os.path.exists(directory):
        os.makedirs(directory)

    # Lưu các hành động vào file JSON
    with open(script_file, 'w') as f:
        json.dump(actions, f, indent=2)

    print(f"Recorded user actions saved to {script_file}")

def replay_user_actions1(script_file, window):
    print(f"Replaying user actions from {script_file}")

    # Đọc file JSON
    with open(script_file, 'r') as f:
        actions = json.load(f)

    start_time = time.time()

    button_map = {
        'Button.left': 'left',
        'Button.middle': 'middle',
        'Button.right': 'right'
    }

    for action in actions:
        # Đợi cho đến khi đến thời điểm thực hiện hành động
        while (time.time() - start_time) < action['time']:
            time.sleep(0.01)

        if action['type'] == 'click':
            # Chuyển đổi tọa độ tương đối thành tọa độ tuyệt đối
            abs_x = window.left + action['x']
            abs_y = window.top + action['y']
            
            # Chuyển đổi giá trị button từ 'Button.left' thành 'left'
            button = button_map.get(action['button'], action['button'])

            # Thực hiện click
            if action['pressed']:
                pyautogui.mouseDown(abs_x, abs_y, button=button)
            else:
                pyautogui.mouseUp(abs_x, abs_y, button=button)

            print(f"Replayed click at ({abs_x}, {abs_y}) with button {button}, pressed: {action['pressed']}")

        elif action['type'] == 'keypress':
            # Thực hiện nhấn phím
            pyautogui.press(action['keys'])
            print(f"Replayed keypress: {action['keys']}")

    print("Replay completed")

def record_user_actions(window, script_file):
    print("Recording user actions. Press 'Esc' to stop recording.")
    actions = []
    start_time = time.time()
    stop_recording = [False]

    def stop():
        stop_recording[0] = True

    kb.Listener.suppress_hotkeys = True

    if window is None:
        print("No active window found.")
        return

    print(f"Recording actions for window: {window.title}")

    def on_click(x, y, button, pressed):
        if window.isActive:
            # Record relative coordinates as percentages
            relative_x = (x - window.left) / window.width
            relative_y = (y - window.top) / window.height
            action = {
                "type": "click",
                "x": relative_x,
                "y": relative_y,
                "time": time.time() - start_time,
                "button": str(button),
                "pressed": pressed
            }
            actions.append(action)
            print(f"Recorded click at ({relative_x:.2f}, {relative_y:.2f}) with button {button}, pressed: {pressed}")

    def on_key_press(key):
        if window.isActive:
            try:
                key_str = key.char
            except AttributeError:
                key_str = str(key)
            action = {
                "type": "keypress",
                "keys": key_str,
                "time": time.time() - start_time
            }
            actions.append(action)
            print(f"Recorded keypress: {key_str}")

    def on_key_release(key):
        if key == kb.Key.esc:
            stop_recording[0] = True
            return False

    mouse_listener = mouse.Listener(on_click=on_click)
    keyboard_listener = kb.Listener(on_press=on_key_press, on_release=on_key_release)

    mouse_listener.start()
    keyboard_listener.start()

    while not stop_recording[0]:
        time.sleep(0.1)

    mouse_listener.stop()
    keyboard_listener.stop()

    directory = os.path.dirname(script_file)
    if not os.path.exists(directory):
        os.makedirs(directory)

    with open(script_file, 'w') as f:
        json.dump(actions, f, indent=2)

    print(f"Recorded user actions saved to {script_file}")

def replay_user_actions(script_file, window):
    print(f"Replaying user actions from {script_file}")

    with open(script_file, 'r') as f:
        actions = json.load(f)

    start_time = time.time()

    button_map = {
        'Button.left': 'left',
        'Button.middle': 'middle',
        'Button.right': 'right'
    }

    for action in actions:
        while (time.time() - start_time) < action['time']:
            time.sleep(0.01)

        if action['type'] == 'click':
            # Convert percentage coordinates to absolute coordinates
            abs_x = int(window.left + (action['x'] * window.width))
            abs_y = int(window.top + (action['y'] * window.height))
            
            button = button_map.get(action['button'], action['button'])

            if action['pressed']:
                pyautogui.mouseDown(abs_x, abs_y, button=button)
            else:
                pyautogui.mouseUp(abs_x, abs_y, button=button)

            print(f"Replayed click at ({abs_x}, {abs_y}) with button {button}, pressed: {action['pressed']}")

        elif action['type'] == 'keypress':
            pyautogui.press(action['keys'])
            print(f"Replayed keypress: {action['keys']}")

    print("Replay completed")

def is_process_running(process_name):
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
            return True
    return False

def find_installation_window(path_data):
    try:
        app = Application(backend='uia').connect(path=path_data)
        logging.info("Installation window found.")
        return app
    except Exception as e:
        logging.error(f"Failed to find installation window: {e}")
        return None

def get_product_name(file_path):
    try:
        pe = pefile.PE(file_path)
        for file_info in pe.FileInfo:
            if file_info[0].Key == b'StringFileInfo':
                for st in file_info[0].StringTable:
                    for entry in st.entries.items():
                        if entry[0] == b'ProductName' or entry[0] == b'productname':
                            return entry[1].decode('utf-8')
    except Exception as e:
        print(f"Error reading product name from {file_path}: {e}")
    return None

def get_active_window_by_pid(pid):
    # Tìm cửa sổ active
    active_window = None
    for window in gw.getAllWindows():
        try:
            if psutil.Process(window.pid).pid == pid:
                active_window = window
                break
        except psutil.NoSuchProcess:
            continue
    
    return active_window

def clean_string(input_string):
    # Loại bỏ khoảng trắng đầu và cuối
    cleaned_string = input_string.strip()
    
    # Loại bỏ dấu ' đầu và cuối
    cleaned_string = re.sub(r"^\s*'|'?\s*$", "", cleaned_string)
    
    return cleaned_string

def install_software(file_path, installed_cache, script_folder):
    if not os.path.exists(file_path):
        logging.error(f"File {file_path} does not exist.")
        return

    file_extension = os.path.splitext(file_path)[1].lower()
    software_name = os.path.basename(file_path).replace(file_extension, '')
    script_file = os.path.join(script_folder, f"{software_name}.js")

    productname = get_product_name(file_path)
    product_name = clean_string(productname)
    # if is_software_installed(software_name, installed_cache,product_name):
    if is_ultraviewer_installed(installed_cache,product_name):
       
        logging.info(f"{software_name} is already installed.")
        return

    if os.path.exists(script_file):
        # logging.info(f"Found script for {software_name}, running script.")
        # run_script(script_file)
        logging.info(f"Found script for {software_name}, replaying actions.")
        process = subprocess.Popen([file_path])
        productname = get_product_name(file_path)
        # product_name = productname.replace("'","")
        product_name = clean_string(productname)
        time.sleep(2)  # Đợi cửa sổ cài đặt xuất hiện
        window = get_window_by_id(product_name , True, process.pid)
        # window = gw.getActiveWindow()
        if window:
            replay_user_actions(script_file, window)
        else:
            logging.error("No active window found for replay.")
        
        # Đợi quá trình cài đặt hoàn tất
        process.wait()
    else:
        try:
            if file_extension == '.exe':
                process = subprocess.Popen([file_path, '/S'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                # process = subprocess.Popen([file_path, '/S'])
                # productname = get_product_name(file_path)
                # product_name = clean_string(productname)
                start_time = time.time()
                max_wait_time = 5  # 60 seconds max wait time
                while process.poll() is None and (time.time() - start_time) < max_wait_time:
                    time.sleep(1)
                if process.poll() is None:
                    logging.warning(f"Silent installation might have failed for {software_name}. Checking for interaction.")
                    try:
                        logging.warning(f"Silent installation failed for {software_name}. Recording user actions.")
                        app = get_window_by_id(product_name , True, process.pid)
                        record_user_actions(app,script_file)
                    except Exception as e:
                        logging.error(f"Error while recording user actions: {e}")
                else:
                    if process.returncode != 0:
                        logging.warning(f"Silent installation failed for {software_name}. Recording user actions.")
                        # record_user_actions(script_file)
                        app = get_window_by_id(product_name , True, process.pid)
                        record_user_actions(app,script_file)
                    else:
                        logging.info(f"Successfully installed {software_name}")
            elif file_extension == '.msi':
                subprocess.run(['msiexec', '/i', file_path, '/quiet', '/norestart'], check=True)
                logging.info(f"Successfully installed {software_name}")
            else:
                logging.warning(f"Unsupported file type: {file_extension}")
        except subprocess.CalledProcessError as e:
            logging.error(f"Failed to install {software_name}: {e}")
            # record_user_actions(script_file)
            app = get_window_by_id(product_name , True, process.pid)
            record_user_actions(app,script_file)

def install_all_softwares_in_folder(folder_path, installed_cache, script_folder):
    if not os.path.exists(folder_path):
        logging.error(f"Folder {folder_path} does not exist.")
        return

    logging.info("Processing software installations...")
    logging.info("=======================================================")
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            install_software(file_path, installed_cache, script_folder)
            logging.info("=======================================================")

if __name__ == "__main__":
    try:
        installed_software_cache = None
        script_folder = ".\\Application\\Scripts"
        path = ".\\Application\\FileSetup"
        installed_cache = is_registry_installed()
        # installed_cache = get_installed_software()
        # installed_cache = []
        # install_all_softwares_in_folder(path, installed_cache, script_folder)
        install_all_softwares_in_folder(path, installed_cache, script_folder)
        logging.info("Process Success...")
    except ValueError as e:
        logging.error(f"Error: {e}")

    input("Press any key to exit...")
