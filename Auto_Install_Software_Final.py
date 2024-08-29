import time
import keyboard
from pynput import mouse
import subprocess
import os
import psutil
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_installed_software():
    global installed_software_cache
    if installed_software_cache is not None:
        return installed_software_cache

    logging.info("Loading installed software data...")
    try:
        # Check for installed software by searching in the list of installed programs
        output = subprocess.check_output(['wmic', 'product', 'get', 'name'], universal_newlines=True)
        installed_software_cache = output.lower().split('\n')
        logging.info("Installed software data loaded successfully.")
    except subprocess.CalledProcessError:
        installed_software_cache = []
        logging.error("Error loading installed software data.")

    return installed_software_cache

def is_software_installed(software_name, installed_cache):
    for software in installed_cache:
        if software_name.lower() in software:
            return True
    return False

def close_installer_window(process_name):
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and process_name.lower() in proc.info['name'].lower():
            proc.terminate()
            try:
                proc.wait(timeout=5)
                logging.info(f"Closed {process_name} installer window.")
            except psutil.TimeoutExpired:
                proc.kill()
                logging.warning(f"Force killed {process_name} installer window.")

def install_software(file_path, installed_cache):
    if not os.path.exists(file_path):
        logging.error(f"File {file_path} does not exist.")
        return

    file_extension = os.path.splitext(file_path)[1].lower()
    software_name = os.path.basename(file_path).replace(file_extension, '')

    if is_software_installed(software_name, installed_cache):
        logging.info(f"{software_name} is already installed.")
        return

    try:
        if file_extension == '.exe':
            # process = subprocess.Popen([file_path, '/S'])
            subprocess.run([file_path, '/S'], check=True)
            logging.info(f"Successfully installed {software_name}")
        elif file_extension == '.msi':
            subprocess.run(['msiexec', '/i', file_path, '/quiet', '/norestart'], check=True)
            logging.info(f"Successfully installed {software_name}")
        else:
            logging.warning(f"Unsupported file type: {file_extension}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to install {software_name}: {e}")

def install_all_softwares_in_folder(folder_path, installed_cache ):
    if not os.path.exists(folder_path):
        logging.error(f"Folder {folder_path} does not exist.")
        return

    logging.info("Processing software installations...")
    logging.info("=======================================================")
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            install_software(file_path, installed_cache)
            logging.info("=======================================================")

def record_user_actions(output_file):
    print("Recording user actions. Press 'q' to stop recording.")
    actions = []
    
    # Define keyboard event to stop recording
    keyboard.add_hotkey('q', lambda: actions.append("EndRecording"))

    # Define mouse event listener
    def on_click(x, y, button, pressed):
        if pressed:
            actions.append(f"Click at ({x}, {y})")

    mouse_listener = mouse.Listener(on_click=on_click)
    mouse_listener.start()

    # Record actions until 'q' is pressed
    while True:
        if "EndRecording" in actions:
            break
        time.sleep(0.1)  # Adjust sleep time based on recording frequency

    mouse_listener.stop()
    keyboard.remove_hotkey('q')

    # Create directory if it doesn't exist
    directory = os.path.dirname(output_file)
    if not os.path.exists(directory):
        os.makedirs(directory)

    # Save recorded actions to file
    with open(output_file, 'w') as f:
        for action in actions:
            f.write(action + '\n')

    print(f"Recorded user actions saved to {output_file}")

if __name__ == "__main__":
    try:
        # Initialize
        installed_software_cache = None
        # file_name_killer = ["readerdc64_en_ka_cra_mdr_install", "SSMS-Setup-ENU"]
        path = ".\\Application"
        path_output_script = ".\\Application\Script"

        installed_cache = get_installed_software()
        if installed_cache is not None:
            install_all_softwares_in_folder(path, installed_cache)
            logging.info("Process Success...")
        else:
            logging.error("Error loading installed software data.")
    except ValueError as e:
        logging.error(f"Error: {e}")

    input("Press any key to exit...")
