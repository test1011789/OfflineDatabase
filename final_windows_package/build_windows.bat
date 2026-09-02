@echo off
setlocal EnableExtensions
chcp 65001 >nul

cd /d "%~dp0"

echo ================================================
echo OfflineDatabase Windows Portable Build
echo ================================================
echo.

echo [1/6] 檢查 Python...
where py >nul 2>&1
if not errorlevel 1 (
    set "PY=py"
    goto :python_ok
)
where python >nul 2>&1
if not errorlevel 1 (
    set "PY=python"
    goto :python_ok
)
echo [錯誤] 找不到 Python。
goto :error

:python_ok
%PY% --version
if errorlevel 1 goto :error

echo.
echo [2/6] 安裝或更新建置工具...
%PY% -m pip install --upgrade openpyxl pyinstaller
if errorlevel 1 goto :error

echo.
echo [3/6] 檢查必要檔案...
if not exist "app.py" (
    echo [錯誤] 找不到 app.py
    goto :error
)
if not exist "initial_data.xlsx" (
    echo [錯誤] 找不到 initial_data.xlsx
    goto :error
)
%PY% -m py_compile app.py
if errorlevel 1 (
    echo [錯誤] app.py 無法通過 Python 語法檢查。
    goto :error
)

echo.
echo [4/6] 清理舊版輸出...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist OfflineDatabase.spec del /q OfflineDatabase.spec

echo.
echo [5/6] 建立 Windows 免安裝程式...
if exist "OfflineDatabase.ico" (
    echo 使用自訂圖示：OfflineDatabase.ico
    %PY% -m PyInstaller --noconfirm --clean --windowed --onedir --name OfflineDatabase --icon OfflineDatabase.ico app.py
) else (
    echo 未找到 OfflineDatabase.ico，使用 PyInstaller 預設圖示。
    %PY% -m PyInstaller --noconfirm --clean --windowed --onedir --name OfflineDatabase app.py
)
if errorlevel 1 goto :error

if not exist "dist\OfflineDatabase\OfflineDatabase.exe" (
    echo [錯誤] PyInstaller 執行完成，但找不到 OfflineDatabase.exe
    goto :error
)

echo.
echo [6/6] 建立 Portable 資料夾結構...
copy /Y "initial_data.xlsx" "dist\OfflineDatabase\initial_data.xlsx" >nul
if errorlevel 1 goto :error

if exist "README_使用說明.md" copy /Y "README_使用說明.md" "dist\OfflineDatabase\README_使用說明.md" >nul
if not exist "dist\OfflineDatabase\data" mkdir "dist\OfflineDatabase\data"
if not exist "dist\OfflineDatabase\backup" mkdir "dist\OfflineDatabase\backup"

if exist "OfflineDatabase.ico" copy /Y "OfflineDatabase.ico" "dist\OfflineDatabase\OfflineDatabase.ico" >nul

echo.
echo ================================================
echo 建置成功！
echo ================================================
echo EXE：%CD%\dist\OfflineDatabase\OfflineDatabase.exe
echo.
echo 使用者端不需要安裝 Python。
echo 請將整個 OfflineDatabase 資料夾一起攜帶。
echo.
pause
exit /b 0

:error
echo.
echo ================================================
echo 建置失敗！
echo ================================================
echo 請保留這個視窗中的錯誤訊息，以便排查。
echo.
pause
exit /b 1
