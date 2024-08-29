import pyautogui
import pygetwindow as gw
import time
import json
import keyboard
from pynput import mouse, keyboard as kb

def record_installation(use_pid=False, pid=None):
    actions = []
    start_time = time.time()
    stop_recording = [False]

    def stop():
        stop_recording[0] = True

    keyboard.add_hotkey('q', stop)

    # Kiểm tra xem có sử dụng PID để lấy cửa sổ hay không
    if use_pid and pid:
        window = None
        windows = gw.getWindowsWithTitle(str(pid))
        if windows:
            window = windows[0]
    else:
        window = gw.getActiveWindow()

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

    # Lưu các hành động vào file JSON
    with open("installation_actions.json", "w") as f:
        json.dump(actions, f, indent=2)

    print("Recording completed. Actions saved to installation_actions.json")
    keyboard.remove_hotkey('q')

# Sử dụng hàm để ghi lại thao tác
record_installation()

# Nếu muốn sử dụng PID, hãy uncomment dòng dưới đây và thay thế PID_NUMBER bằng PID thực tế
# record_installation(use_pid=True, pid=PID_NUMBER)
