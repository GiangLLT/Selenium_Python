@echo off
REM Di chuyển đến thư mục của dự án
cd /d %~dp0

REM Kiểm tra xem Python đã được cài đặt chưa
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Python is not installed. Please install Python and try again.
    pause
    exit /b
)

REM Kiểm tra xem tệp Surver_Selenium.py có tồn tại không
IF NOT EXIST Surver_Selenium.py (
    echo Surver_Selenium.py not found. Please ensure the file exists in the project directory.
    pause
    exit /b
)

REM Thực thi file Surver_Selenium.py bằng Python
echo Running Surver_Selenium.py...
python Surver_Selenium.py

REM Thông báo hoàn thành
echo Surver_Selenium.py has been executed successfully.
pause
