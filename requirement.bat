@echo off
REM Di chuyển đến thư mục của dự án
cd /d %~dp0

REM Kiểm tra xem Python đã được cài đặt chưa
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Python is not installed. Installing Python now...
    REM Thực hiện cài đặt Python tự động
    powershell.exe -Command "Invoke-WebRequest -Uri https://www.python.org/ftp/python/3.11.2/python-3.11.2-amd64.exe -OutFile python-3.11.2-amd64.exe"
    python-3.11.2-amd64.exe /quiet InstallAllUsers=1 PrependPath=1 Include_test=0
    IF %ERRORLEVEL% NEQ 0 (
        echo Failed to install Python. Please install Python manually and try again.
        pause
        exit /b
    )
)

REM Kiểm tra xem tệp requirements.txt có tồn tại không
IF NOT EXIST requirements.txt (
    echo requirements.txt not found. Please ensure the file exists in the project directory.
    pause
    exit /b
)

REM Cài đặt các thư viện từ tệp requirements.txt
echo Installing required libraries from requirements.txt...
pip install -r requirements.txt

REM Thông báo hoàn thành
echo All libraries have been installed successfully.
pause
