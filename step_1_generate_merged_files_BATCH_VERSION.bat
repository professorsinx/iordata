@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM                    BATCH FILE VERSION OF STEP_1 SCRIPT
REM ============================================================================
REM Purpose: Merges data from BAL, ENT, and IOR CSV files
REM Author: Created as batch file alternative to PowerShell script
REM Date: %date%
REM
REM IMPORTANT BEGINNER NOTES:
REM 1. This batch file requires Python to be installed on your system
REM 2. Make sure Python is in your system PATH
REM 3. The script will check for Python availability before proceeding
REM 4. All CSV files must be properly formatted (comma-separated)
REM 5. File names must match the expected patterns (BAL*.csv, ENT*.csv, IOR*.csv)
REM
REM BEFORE USING THIS SCRIPT:
REM - Ensure input folder exists with BAL*.csv and ENT*.csv files
REM - Ensure IOR*.csv file exists in the same folder as this batch file
REM - Make sure you have write permissions to the output folder
REM - Close any Excel files that might be using the output files
REM ============================================================================

echo.
echo ========================================
echo   CSV DATA MERGER - BATCH VERSION
echo ========================================
echo.

REM [1/9] INITIALIZING SCRIPT
echo [1/9] INITIALIZING SCRIPT...

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH!
    echo Please install Python from https://python.org
    echo Make sure to check "Add Python to PATH" during installation
    goto :error_exit
)

REM Create output directory if it doesn't exist
if not exist "output" (
    echo   Creating output directory...
    mkdir output
)

REM Store current date for calculations
for /f "tokens=2 delims==" %%a in ('wmic OS Get localdatetime /value') do set "dt=%%a"
set "current_date=!dt:~0,4!-!dt:~4,2!-!dt:~6,2!"

echo   Current date: !current_date!
echo   Script initialized successfully.

REM [2/9] LOCATING FILES
echo.
echo [2/9] LOCATING FILES...

REM Find BAL file in input directory
set "bal_file="
for %%f in (input\BAL*.csv) do (
    set "bal_file=%%f"
    goto :bal_found
)
:bal_found

REM Find ENT file in input directory  
set "ent_file="
for %%f in (input\ENT*.csv) do (
    set "ent_file=%%f"
    goto :ent_found
)
:ent_found

REM Find IOR file in current directory
set "ior_file="
for %%f in (IOR*.csv) do (
    set "ior_file=%%f"
    goto :ior_found
)
:ior_found

REM Validate all files exist
set "missing_files="
if "!bal_file!"=="" set "missing_files=!missing_files! BAL*.csv"
if "!ent_file!"=="" set "missing_files=!missing_files! ENT*.csv"  
if "!ior_file!"=="" set "missing_files=!missing_files! IOR*.csv"

if not "!missing_files!"=="" (
    echo [ERROR] Missing required files:!missing_files!
    echo.
    echo Please ensure the following files exist:
    echo - BAL*.csv file in the 'input' folder
    echo - ENT*.csv file in the 'input' folder  
    echo - IOR*.csv file in the current folder
    goto :error_exit
)

echo   Found files:
echo     BAL File: !bal_file!
echo     ENT File: !ent_file!
echo     IOR File: !ior_file!
echo     Output: output\Merged_Output.csv

REM [3/9] PROCESSING DATA WITH PYTHON
echo.
echo [3/9] PROCESSING DATA WITH PYTHON...
echo   Creating Python processing script...

REM Create temporary Python script to handle the complex data processing
(
echo import pandas as pd
echo import sys
echo from datetime import datetime
echo import os
echo.
echo def main^(^):
echo     try:
echo         print^("  Loading BAL data..."^)
echo         # Read BAL file and filter
echo         bal_data = pd.read_csv^('!bal_file!'^ 
echo         bal_filtered = bal_data[^(bal_data['MODULE'] == 'TRICHY'^) ^&
echo                                ^(bal_data['BGL'].isin^(['98556', '98585', '98641']^)^)^]
echo.
echo         # Add required columns to BAL data
echo         bal_filtered['REF'^] = 'X'
echo         bal_filtered['AMOUNT'^] = bal_filtered['BALANCE'^]
echo         bal_filtered['POST DATE'^] = 'X'
echo         bal_filtered['DAYS'^] = 'X'
echo.
echo         print^(f"  BAL records after filtering: {len^(bal_filtered^)}"^)
echo.
echo         print^("  Loading ENT data..."^)
echo         # Read ENT file and filter
echo         ent_data = pd.read_csv^('!ent_file!'^ 
echo         ent_filtered = ent_data[ent_data['MODULE'^] == 'TRICHY'^]
echo.
echo         print^(f"  ENT records after filtering: {len^(ent_filtered^)}"^)
echo.
echo         print^("  Combining BAL and ENT data..."^)
echo         # Combine BAL and ENT data
echo         combined_data = pd.concat^([bal_filtered, ent_filtered^], ignore_index=True^)
echo.
echo         print^(f"  Total combined records: {len^(combined_data^)}"^)
echo.
echo         print^("  Loading IOR data..."^)
echo         # Read IOR file ^(skip first row, add headers^)
echo         ior_data = pd.read_csv^('!ior_file!', skiprows=1, 
echo                                names=['SLNO_A', 'BGL', 'BGL_NAME', 'TAT', 'SLNO_B'^]^)
echo         ior_data = ior_data[['BGL', 'TAT', 'SLNO_B'^]^].rename^(columns={'SLNO_B': 'SL NO'}^)
echo.
echo         print^(f"  IOR records processed: {len^(ior_data^)}"^)
echo.
echo         print^("  Performing initial merge..."^)
echo         # Merge combined data with IOR data
echo         merged_data = pd.merge^(combined_data, ior_data, on='BGL', how='left'^)
echo.
echo         # Fill missing SL NO with 9
echo         merged_data['SL NO'^].fillna^(9, inplace=True^)
echo.
echo         # Add missing columns
echo         merged_data['OVERDUE'^] = None
echo         merged_data['DAYS PASSED'^] = None
echo.
echo         print^("  Applying business rules..."^)
echo         # Apply business rules for specific BGLs
echo         for idx, row in merged_data.iterrows^(^):
echo             bgl = str^(row['BGL'^]^).strip^(^)
echo             amount = 0
echo             
echo             try:
echo                 if pd.notna^(row['AMOUNT'^]^):
echo                     amount = abs^(float^(str^(row['AMOUNT'^]^).replace^(',', '^'^).replace^('$', '^'^)^)^)
echo             except:
echo                 amount = 0
echo.
echo             # Apply BGL-specific rules
echo             if bgl == '98533':
echo                 merged_data.at[idx, 'SL NO'^] = 4 if amount ^>= 10000000 else 5
echo             elif bgl == '2399869':
echo                 merged_data.at[idx, 'SL NO'^] = 4 if amount ^>= 10000000 else 6
echo             elif bgl == '98593':
echo                 merged_data.at[idx, 'SL NO'^] = 4 if amount ^>= 10000000 else 5
echo.
echo         print^("  Performing final calculations..."^)
echo         # Calculate DAYS PASSED and OVERDUE
echo         current_date = datetime.now^(^)
echo         
echo         for idx, row in merged_data.iterrows^(^):
echo             post_date_str = str^(row['POST DATE'^]^).strip^(^)
echo             
echo             if post_date_str != 'X' and post_date_str != 'nan':
echo                 try:
echo                     post_date = datetime.strptime^(post_date_str, '%%Y-%%m-%%d'^)
echo                     days_passed = ^(current_date - post_date^).days
echo                     merged_data.at[idx, 'DAYS PASSED'^] = days_passed
echo                     
echo                     # Calculate OVERDUE
echo                     tat = row['TAT'^]
echo                     if pd.notna^(tat^):
echo                         try:
echo                             tat_days = int^(str^(tat^).replace^('D', '^'^).strip^(^)^)
echo                             merged_data.at[idx, 'OVERDUE'^] = days_passed - tat_days
echo                         except:
echo                             merged_data.at[idx, 'OVERDUE'^] = 'NA'
echo                     else:
echo                         merged_data.at[idx, 'OVERDUE'^] = 'NA'
echo                 except:
echo                     merged_data.at[idx, 'DAYS PASSED'^] = 0
echo                     merged_data.at[idx, 'OVERDUE'^] = 'NA'
echo             else:
echo                 merged_data.at[idx, 'DAYS PASSED'^] = 0
echo                 merged_data.at[idx, 'OVERDUE'^] = 'NA'
echo.
echo         print^("  Exporting results..."^)
echo         # Export to CSV
echo         merged_data.to_csv^('output/Merged_Output.csv', index=False^)
echo         
echo         print^(f"  Successfully processed {len^(merged_data^)} records"^)
echo         print^("  Output saved to: output/Merged_Output.csv"^)
echo         
echo         return 0
echo.
echo     except Exception as e:
echo         print^(f"[ERROR] {str^(e^)}"^)
echo         return 1
echo.
echo if __name__ == "__main__":
echo     sys.exit^(main^(^)^)
) > temp_processor.py

REM [4/9] RUNNING PYTHON PROCESSOR
echo.
echo [4/9] RUNNING PYTHON PROCESSOR...

python temp_processor.py
set "python_result=!errorlevel!"

REM [5/9] CLEANING UP
echo.
echo [5/9] CLEANING UP...
del temp_processor.py

if !python_result! neq 0 (
    echo [ERROR] Python processing failed!
    goto :error_exit
)

REM [6/9] VERIFICATION
echo.
echo [6/9] VERIFYING OUTPUT...

if not exist "output\Merged_Output.csv" (
    echo [ERROR] Output file was not created!
    goto :error_exit
)

REM Count lines in output file (subtract 1 for header)
for /f %%a in ('type "output\Merged_Output.csv" ^| find /c /v ""') do set "line_count=%%a"
set /a "record_count=!line_count!-1"

echo   Output file exists: output\Merged_Output.csv
echo   Total records processed: !record_count!

REM [7/9] SCRIPT COMPLETED
echo.
echo ========================================
echo   SCRIPT COMPLETED SUCCESSFULLY!
echo ========================================
echo   Output Directory: output
echo   Output File: Merged_Output.csv
echo   Total Records: !record_count!
echo   Completion Time: %time%
echo ========================================
echo.

goto :normal_exit

:error_exit
echo.
echo ========================================
echo   SCRIPT FAILED!
echo ========================================
echo.
echo TROUBLESHOOTING TIPS:
echo 1. Check that all required CSV files exist
echo 2. Ensure Python is installed and in PATH
echo 3. Verify CSV files are not corrupted
echo 4. Make sure you have write permissions
echo 5. Close any Excel files using the output
echo.
pause
exit /b 1

:normal_exit
echo Press any key to exit...
pause >nul
exit /b 0