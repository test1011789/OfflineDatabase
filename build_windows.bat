@echo off
setlocal EnableExtensions

cd /d "%~dp0"

echo ==========================================
echo OfflineDatabase Windows Portable Builder
echo ==========================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found.
    echo This script is intended to run from GitHub Actions or a Windows build environment with Python.
    goto :error
)

echo [1/5] Checking Python...
python --version
if errorlevel 1 goto :error

echo.
echo [2/5] Checking build dependencies...
python -c "import openpyxl, PyInstaller" >nul 2>nul
if errorlevel 1 (
    echo Installing openpyxl and PyInstaller...
    python -m pip install --upgrade openpyxl pyinstaller
    if errorlevel 1 goto :error
)

echo.
echo [3/5] Cleaning old build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist OfflineDatabase.spec del /q OfflineDatabase.spec

if not exist app.py (
    echo [ERROR] app.py was not found.
    goto :error
)

if not exist initial_data.xlsx (
    echo [ERROR] initial_data.xlsx was not found.
    goto :error
)

echo.
echo [4/5] Building Windows Portable EXE...
set "ICON_ARG="
if exist OfflineDatabase.ico (
    echo Custom icon found: OfflineDatabase.ico
    set "ICON_ARG=--icon OfflineDatabase.ico"
) else (
    echo No OfflineDatabase.ico found. Building with the default icon.
)

python -m PyInstaller --noconfirm --clean --windowed --onedir --name OfflineDatabase %ICON_ARG% app.py
if errorlevel 1 goto :error

echo.
echo [5/5] Preparing Portable folder...
copy /Y "initial_data.xlsx" "dist\OfflineDatabase\initial_data.xlsx" >nul
if errorlevel 1 goto :error

if not exist "dist\OfflineDatabase\data" mkdir "dist\OfflineDatabase\data"
if not exist "dist\OfflineDatabase\backup" mkdir "dist\OfflineDatabase\backup"

if exist "README_使用說明.md" copy /Y "README_使用說明.md" "dist\OfflineDatabase\README_使用說明.md" >nul
if exist "README_Σ╜┐τö¿Φ¬¬µÿÄ.md" copy /Y "README_Σ╜┐τö¿Φ¬¬µÿÄ.md" "dist\OfflineDatabase\README_使用說明.md" >nul

if exist "OfflineDatabase.ico" copy /Y "OfflineDatabase.ico" "dist\OfflineDatabase\OfflineDatabase.ico" >nul

if not exist "dist\OfflineDatabase\OfflineDatabase.exe" (
    echo [ERROR] OfflineDatabase.exe was not created.
    goto :error
)

echo.
echo ==========================================
echo BUILD SUCCESS
echo ==========================================
echo EXE:
echo %CD%\dist\OfflineDatabase\OfflineDatabase.exe
echo.
echo Portable folder:
echo %CD%\dist\OfflineDatabase\
echo.
exit /b 0

:error
echo.
echo ==========================================
echo BUILD FAILED
echo ==========================================
echo Please read the error message above.
echo.
exit /b 1
