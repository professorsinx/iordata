@echo off
setlocal enabledelayedexpansion

:: Define the output directory
set "outputDir=%~dp0output"

:: Check if output directory exists
if not exist "%outputDir%" (
    echo Output directory does not exist!
    exit /b
)

:: Run the VBScript
cscript //nologo "%~dp0merge_csv_to_excel.vbs" "%outputDir%"

echo Excel file has been created successfully in the output folder.
pause
