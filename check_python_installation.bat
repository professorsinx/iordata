@echo off
echo ========================================
echo    PYTHON INSTALLATION CHECKER
echo ========================================
echo.

REM Method 1: Check if python command is available
echo [Method 1] Checking if 'python' command is available...
python --version >nul 2>&1
if %errorlevel% equ 0 (
    echo ✓ Python is installed and accessible via 'python' command
    python --version
) else (
    echo ✗ Python 'python' command not found or not in PATH
)

echo.

REM Method 2: Check if py launcher is available (Windows Python Launcher)
echo [Method 2] Checking if 'py' launcher is available...
py --version >nul 2>&1
if %errorlevel% equ 0 (
    echo ✓ Python is installed and accessible via 'py' launcher
    py --version
) else (
    echo ✗ Python 'py' launcher not found
)

echo.

REM Method 3: Check if python3 command is available
echo [Method 3] Checking if 'python3' command is available...
python3 --version >nul 2>&1
if %errorlevel% equ 0 (
    echo ✓ Python is installed and accessible via 'python3' command
    python3 --version
) else (
    echo ✗ Python 'python3' command not found
)

echo.

REM Method 4: Check common installation directories
echo [Method 4] Checking common Python installation directories...

set "python_found=0"

REM Check Program Files
if exist "C:\Program Files\Python*" (
    echo ✓ Found Python in C:\Program Files\
    dir "C:\Program Files\Python*" /b
    set "python_found=1"
)

REM Check Program Files (x86)
if exist "C:\Program Files (x86)\Python*" (
    echo ✓ Found Python in C:\Program Files (x86)\
    dir "C:\Program Files (x86)\Python*" /b
    set "python_found=1"
)

REM Check AppData Local
if exist "%LOCALAPPDATA%\Programs\Python" (
    echo ✓ Found Python in %LOCALAPPDATA%\Programs\Python\
    dir "%LOCALAPPDATA%\Programs\Python" /b
    set "python_found=1"
)

REM Check Users directory
if exist "C:\Users\%USERNAME%\AppData\Local\Programs\Python" (
    echo ✓ Found Python in user-specific location
    set "python_found=1"
)

if %python_found% equ 0 (
    echo ✗ No Python installations found in common directories
)

echo.

REM Method 5: Check Windows Registry for Python installations
echo [Method 5] Checking Windows Registry for Python...
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Python" >nul 2>&1
if %errorlevel% equ 0 (
    echo ✓ Python registry entries found
    reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Python" /s | findstr /i "python"
) else (
    echo ✗ No Python registry entries found
)

echo.

REM Method 6: Check PATH environment variable
echo [Method 6] Checking if Python is in PATH environment variable...
echo %PATH% | findstr /i python >nul
if %errorlevel% equ 0 (
    echo ✓ Python found in PATH environment variable
    echo Python paths in PATH:
    for %%i in ("%PATH:;=";"%") do (
        echo %%i | findstr /i python >nul
        if not errorlevel 1 echo   %%~i
    )
) else (
    echo ✗ Python not found in PATH environment variable
)

echo.

REM Summary
echo ========================================
echo                SUMMARY
echo ========================================

REM Final check with detailed output
python --version >temp_check.txt 2>&1
if %errorlevel% equ 0 (
    echo ✓ PYTHON IS INSTALLED AND READY TO USE
    echo   Version: 
    type temp_check.txt
    echo   You can use Python in your batch scripts!
) else (
    py --version >temp_check.txt 2>&1
    if !errorlevel! equ 0 (
        echo ✓ PYTHON IS INSTALLED (use 'py' command)
        echo   Version: 
        type temp_check.txt
        echo   Use 'py' instead of 'python' in your scripts
    ) else (
        echo ✗ PYTHON IS NOT ACCESSIBLE
        echo   This means either:
        echo   - Python is not installed
        echo   - Python is installed but not in PATH
        echo   - Installation is corrupted
    )
)

REM Cleanup
if exist temp_check.txt del temp_check.txt

echo.
echo ========================================
echo Press any key to exit...
pause >nul