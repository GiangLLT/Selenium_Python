import subprocess
import os
import psutil
import time

def get_installed_software():
    global installed_software_cache
    if installed_software_cache is not None:
        return installed_software_cache

    print("Loading installed software data...")
    try:
        # Check for installed software by searching in the list of installed programs
        output = subprocess.check_output(['wmic', 'product', 'get', 'name'], universal_newlines=True)
        installed_software_cache = output.lower().split('\n')
        print("Installed software data loaded successfully.")
    except subprocess.CalledProcessError:
        installed_software_cache = []
        print("Error loading installed software data.")

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
                print(f"Closed {process_name} installer window.")
            except psutil.TimeoutExpired:
                proc.kill()
                print(f"Force killed {process_name} installer window.")

def install_software(file_path, installed_cache):
    if not os.path.exists(file_path):
        print(f"File {file_path} does not exist.")
        return

    file_extension = os.path.splitext(file_path)[1].lower()
    software_name = os.path.basename(file_path).replace(file_extension, '')

    if is_software_installed(software_name, installed_cache):
        print(f"{software_name} is already installed.")
        return

    try:
        if file_extension == '.exe':
            subprocess.run([file_path, '/S'], check=True)
            print(f"Successfully installed {software_name}")
        elif file_extension == '.msi':
            subprocess.run(['msiexec', '/i', file_path, '/quiet', '/norestart'], check=True)
            print(f"Successfully installed {software_name}")
        else:
            print(f"Unsupported file type: {file_extension}")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install {software_name}: {e}")

def install_all_softwares_in_folder(folder_path, installed_cache):
    if not os.path.exists(folder_path):
        print(f"Folder {folder_path} does not exist.")
        return

    print("Processing software installations...")
    print("=======================================================")
    for file_name in os.listdir(folder_path):
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            install_software(file_path, installed_cache)
            print("=======================================================")

if __name__ == "__main__":
    try:
        #khởi tạo
        installed_software_cache = None
        # file_name_killer = ["readerdc64_en_ka_cra_mdr_install", "SSMS-Setup-ENU"]
        path = ".\\Application"

        installed_cache = get_installed_software()
        if installed_cache is not None:
            install_all_softwares_in_folder(path, installed_cache)
            print("Process Success...")
        else:
            print("Error loading installed software data.")
    except ValueError as e:
        print(f"Error: {e}")

    input("Press any key to exit...")






# import subprocess
# import os

# installed_software_cache = None

# def get_installed_software():
#     print("Load data software installed ....")
#     global installed_software_cache
#     if installed_software_cache is not None:
#         return installed_software_cache

#     try:
#         # Check for installed software by searching in the list of installed programs
#         output = subprocess.check_output(['wmic', 'product', 'get', 'name'], universal_newlines=True)
#         print("Load data software successs ....")
#         installed_software_cache = output.lower().split('\n')
#     except subprocess.CalledProcessError:
#         installed_software_cache = []

#     return installed_software_cache

# def is_software_installed(software_name):
#     installed_software = get_installed_software()
#     for software in installed_software:
#         if software_name.lower() in software:
#             return True
#     return False

# def install_software(file_path):
#     if not os.path.exists(file_path):
#         print(f"File {file_path} does not exist.")
#         return

#     file_extension = os.path.splitext(file_path)[1].lower()
#     software_name = os.path.basename(file_path).replace(file_extension, '')

#     if is_software_installed(software_name):
#         print(f"{software_name} is already installed.")
#         return

#     try:
#         if file_extension == '.exe':
#             subprocess.run([file_path, '/S'], check=True)
#             print(f"Successfully installed {software_name}")
#         elif file_extension == '.msi':
#             subprocess.run(['msiexec', '/i', file_path, '/quiet', '/norestart'], check=True)
#             print(f"Successfully installed {software_name}")
#         else:
#             print(f"Unsupported file type: {file_extension}")
#     except subprocess.CalledProcessError as e:
#         print(f"Failed to install {software_name}: {e}")

# def install_all_softwares_in_folder(folder_path):
#     if not os.path.exists(folder_path):
#         print(f"Folder {folder_path} does not exist.")
#         return
    
#     print("Sofware install Processing...")
#     print("=======================================================.")
#     for file_name in os.listdir(folder_path):
#         file_path = os.path.join(folder_path, file_name)
#         if os.path.isfile(file_path):
#             install_software(file_path)
#             print("=======================================================.")          

# if __name__ == "__main__":
#     try:
#         path = ".\Application"

#         install_all_softwares_in_folder(path)
#     except ValueError:
#         print("Error")

#     input("Nhấn phím bất kỳ để kết thúc...")  
