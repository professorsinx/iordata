@echo off

REM #########################################################

REM Batch Script: Convert Excel (.xlsx) to CSV without PowerShell

REM #########################################################


REM Step 1: Define input directory

SET "INPUT_DIR=%~dp0input"


REM Check if input directory exists

if not exist "%INPUT_DIR%" (

    echo [ERROR] Input directory not found: %INPUT_DIR%

    pause

    exit /b 1

)


REM Step 2: Search for files starting with "BAL" or "ENT" in the input directory

for %%F in ("%INPUT_DIR%\BAL*.xlsx") do (

    echo [INFO] Converting %%~nxF to CSV...

    cscript //nologo "%~dp0convert_xlsx_to_csv.vbs" "%%F" "%INPUT_DIR%\%%~nF.csv"

)


for %%F in ("%INPUT_DIR%\ENT*.xlsx") do (

    echo [INFO] Converting %%~nxF to CSV...

    cscript //nologo "%~dp0convert_xlsx_to_csv.vbs" "%%F" "%INPUT_DIR%\%%~nF.csv"

)


REM Step 3: Completion message

echo [INFO] Excel-to-CSV conversion process completed.

pause


