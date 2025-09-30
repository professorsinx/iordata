@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM                    PURE BATCH + VBSCRIPT VERSION OF STEP_1 
REM ============================================================================
REM Purpose: Merges data from BAL, ENT, and IOR CSV files using only Windows built-ins
REM Requirements: Windows with VBScript support (available on all Windows systems)
REM No external dependencies required!
REM ============================================================================

echo.
echo ========================================
echo   CSV DATA MERGER - PURE BATCH VERSION
echo ========================================
echo Using VBScript for CSV processing (no Python required)
echo.

REM [1/9] INITIALIZING SCRIPT
echo [1/9] INITIALIZING SCRIPT...

REM Create output directory if it doesn't exist
if not exist "output" (
    echo   Creating output directory...
    mkdir output
)

REM Get current date
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

REM [3/9] CREATING VBSCRIPT PROCESSOR
echo.
echo [3/9] CREATING VBSCRIPT PROCESSOR...

REM Create VBScript to handle CSV processing
(
echo Option Explicit
echo.
echo Dim fso, balFile, entFile, iorFile, outputFile
echo Dim balData, entData, iorData, mergedData
echo Dim i, j, line, fields, record
echo Dim currentDate, postDate, daysPassed, tatValue, overdue
echo.
echo Set fso = CreateObject^("Scripting.FileSystemObject"^)
echo currentDate = Date
echo.
echo ' Function to parse CSV line respecting quoted fields
echo Function ParseCSVLine^(csvLine^)
echo     Dim result^(^), inQuotes, field, char, i
echo     ReDim result^(100^)  ' Initial size
echo     Dim fieldCount, currentField
echo     fieldCount = 0
echo     currentField = ""
echo     inQuotes = False
echo     
echo     For i = 1 To Len^(csvLine^)
echo         char = Mid^(csvLine, i, 1^)
echo         If char = """" Then
echo             inQuotes = Not inQuotes
echo         ElseIf char = "," And Not inQuotes Then
echo             result^(fieldCount^) = Trim^(currentField^)
echo             fieldCount = fieldCount + 1
echo             currentField = ""
echo         Else
echo             currentField = currentField + char
echo         End If
echo     Next
echo     
echo     result^(fieldCount^) = Trim^(currentField^)
echo     ReDim Preserve result^(fieldCount^)
echo     ParseCSVLine = result
echo End Function
echo.
echo ' Function to safely convert amount to number
echo Function SafeAmount^(amountStr^)
echo     Dim cleanAmount
echo     If IsEmpty^(amountStr^) Or amountStr = "" Or amountStr = "X" Then
echo         SafeAmount = 0
echo         Exit Function
echo     End If
echo     
echo     cleanAmount = Replace^(Replace^(Replace^(amountStr, ",", ""^), "$", ""^), " ", ""^)
echo     If IsNumeric^(cleanAmount^) Then
echo         SafeAmount = Abs^(CDbl^(cleanAmount^)^)
echo     Else
echo         SafeAmount = 0
echo     End If
echo End Function
echo.
echo ' Function to determine SL NO based on BGL and amount
echo Function GetSlNo^(bgl, amount^)
echo     Dim amountVal
echo     amountVal = SafeAmount^(amount^)
echo     
echo     Select Case Trim^(bgl^)
echo         Case "98533"
echo             If amountVal ^>= 10000000 Then
echo                 GetSlNo = 4
echo             Else
echo                 GetSlNo = 5
echo             End If
echo         Case "2399869"
echo             If amountVal ^>= 10000000 Then
echo                 GetSlNo = 4
echo             Else
echo                 GetSlNo = 6
echo             End If
echo         Case "98593"
echo             If amountVal ^>= 10000000 Then
echo                 GetSlNo = 4
echo             Else
echo                 GetSlNo = 5
echo             End If
echo         Case Else
echo             GetSlNo = 9  ' Default value
echo     End Select
echo End Function
echo.
echo ' Main processing
echo WScript.Echo "  Loading and processing BAL data..."
echo.
echo ' Read BAL file
echo Set balFile = fso.OpenTextFile^("!bal_file!", 1^)
echo Dim balHeaders, balLines^(^), balCount
echo balHeaders = balFile.ReadLine  ' Read header
echo balCount = 0
echo ReDim balLines^(1000^)  ' Initial size
echo.
echo Do While Not balFile.AtEndOfStream
echo     line = balFile.ReadLine
echo     If Len^(Trim^(line^)^) ^> 0 Then
echo         fields = ParseCSVLine^(line^)
echo         ' Check if MODULE=TRICHY and BGL in list
echo         If UBound^(fields^) ^>= 5 Then  ' Ensure enough fields
echo             If Trim^(fields^(5^)^) = "TRICHY" And ^(Trim^(fields^(7^)^) = "98556" Or Trim^(fields^(7^)^) = "98585" Or Trim^(fields^(7^)^) = "98641"^) Then
echo                 Set record = CreateObject^("Scripting.Dictionary"^)
echo                 record.Add "BRCD", fields^(0^)
echo                 record.Add "BRANCH", fields^(1^)
echo                 record.Add "CIRCLE", fields^(2^)
echo                 record.Add "NETWORK", fields^(3^)
echo                 record.Add "MOD_NO", fields^(4^)
echo                 record.Add "MODULE", fields^(5^)
echo                 record.Add "REG_NO", fields^(6^)
echo                 record.Add "BGL", fields^(7^)
echo                 record.Add "BGLDESC", fields^(8^)
echo                 record.Add "REF", "X"
echo                 record.Add "AMOUNT", fields^(9^)  ' BALANCE field
echo                 record.Add "POST_DATE", "X"
echo                 record.Add "DAYS", "X"
echo                 record.Add "SOURCE", "BAL"
echo                 
echo                 Set balLines^(balCount^) = record
echo                 balCount = balCount + 1
echo                 If balCount Mod 100 = 0 Then ReDim Preserve balLines^(balCount + 100^)
echo             End If
echo         End If
echo     End If
echo Loop
echo balFile.Close
echo ReDim Preserve balLines^(balCount - 1^)
echo.
echo WScript.Echo "  BAL records after filtering: " ^& balCount
echo.
echo ' Read ENT file
echo WScript.Echo "  Loading and processing ENT data..."
echo Set entFile = fso.OpenTextFile^("!ent_file!", 1^)
echo Dim entHeaders, entLines^(^), entCount
echo entHeaders = entFile.ReadLine  ' Read header
echo entCount = 0
echo ReDim entLines^(1000^)  ' Initial size
echo.
echo Do While Not entFile.AtEndOfStream
echo     line = entFile.ReadLine
echo     If Len^(Trim^(line^)^) ^> 0 Then
echo         fields = ParseCSVLine^(line^)
echo         ' Check if MODULE=TRICHY
echo         If UBound^(fields^) ^>= 5 Then  ' Ensure enough fields
echo             If Trim^(fields^(5^)^) = "TRICHY" Then
echo                 Set record = CreateObject^("Scripting.Dictionary"^)
echo                 record.Add "BRCD", fields^(0^)
echo                 record.Add "BRANCH", fields^(1^)
echo                 record.Add "CIRCLE", fields^(2^)
echo                 record.Add "NETWORK", fields^(3^)
echo                 record.Add "MOD_NO", fields^(4^)
echo                 record.Add "MODULE", fields^(5^)
echo                 record.Add "REG_NO", fields^(6^)
echo                 record.Add "BGL", fields^(7^)
echo                 record.Add "BGLDESC", fields^(8^)
echo                 record.Add "REF", fields^(9^)
echo                 record.Add "AMOUNT", fields^(10^)
echo                 record.Add "POST_DATE", fields^(11^)
echo                 record.Add "DAYS", fields^(12^)
echo                 record.Add "SOURCE", "ENT"
echo                 
echo                 Set entLines^(entCount^) = record
echo                 entCount = entCount + 1
echo                 If entCount Mod 100 = 0 Then ReDim Preserve entLines^(entCount + 100^)
echo             End If
echo         End If
echo     End If
echo Loop
echo entFile.Close
echo ReDim Preserve entLines^(entCount - 1^)
echo.
echo WScript.Echo "  ENT records after filtering: " ^& entCount
echo.
echo ' Read IOR file
echo WScript.Echo "  Loading IOR data..."
echo Set iorFile = fso.OpenTextFile^("!ior_file!", 1^)
echo Dim iorHeaders, iorDict
echo iorHeaders = iorFile.ReadLine  ' Skip first line
echo Set iorDict = CreateObject^("Scripting.Dictionary"^)
echo.
echo Do While Not iorFile.AtEndOfStream
echo     line = iorFile.ReadLine
echo     If Len^(Trim^(line^)^) ^> 0 Then
echo         fields = ParseCSVLine^(line^)
echo         If UBound^(fields^) ^>= 4 Then
echo             iorDict.Add Trim^(fields^(1^)^), Array^(Trim^(fields^(3^)^), Trim^(fields^(4^)^)^)  ' BGL -> [TAT, SL_NO]
echo         End If
echo     End If
echo Loop
echo iorFile.Close
echo.
echo WScript.Echo "  IOR records processed: " ^& iorDict.Count
echo.
echo ' Combine and process data
echo WScript.Echo "  Combining data and applying business rules..."
echo.
echo Set outputFile = fso.CreateTextFile^("output\Merged_Output.csv", True^)
echo outputFile.WriteLine "BRCD,BRANCH,CIRCLE,NETWORK,MOD NO,MODULE,REG NO,BGL,BGLDESC,REF,AMOUNT,POST DATE,DAYS,TAT,SL NO,OVERDUE,DAYS PASSED"
echo.
echo Dim totalRecords, processedCount
echo totalRecords = balCount + entCount
echo processedCount = 0
echo.
echo ' Process BAL records
echo For i = 0 To balCount - 1
echo     Dim bgl, amount, slNo, tat
echo     bgl = balLines^(i^).Item^("BGL"^)
echo     amount = balLines^(i^).Item^("AMOUNT"^)
echo     
echo     ' Get SL NO from business rules
echo     slNo = GetSlNo^(bgl, amount^)
echo     
echo     ' Get TAT from IOR data
echo     tat = ""
echo     If iorDict.Exists^(bgl^) Then
echo         tat = iorDict^(bgl^)^(0^)
echo     End If
echo     
echo     ' Write record
echo     outputFile.WriteLine balLines^(i^).Item^("BRCD"^) ^& "," ^& _
echo                         balLines^(i^).Item^("BRANCH"^) ^& "," ^& _
echo                         balLines^(i^).Item^("CIRCLE"^) ^& "," ^& _
echo                         balLines^(i^).Item^("NETWORK"^) ^& "," ^& _
echo                         balLines^(i^).Item^("MOD_NO"^) ^& "," ^& _
echo                         balLines^(i^).Item^("MODULE"^) ^& "," ^& _
echo                         balLines^(i^).Item^("REG_NO"^) ^& "," ^& _
echo                         balLines^(i^).Item^("BGL"^) ^& "," ^& _
echo                         balLines^(i^).Item^("BGLDESC"^) ^& "," ^& _
echo                         balLines^(i^).Item^("REF"^) ^& "," ^& _
echo                         balLines^(i^).Item^("AMOUNT"^) ^& "," ^& _
echo                         balLines^(i^).Item^("POST_DATE"^) ^& "," ^& _
echo                         balLines^(i^).Item^("DAYS"^) ^& "," ^& _
echo                         tat ^& "," ^& _
echo                         slNo ^& "," ^& _
echo                         "NA," ^& _
echo                         "0"
echo     
echo     processedCount = processedCount + 1
echo Next
echo.
echo ' Process ENT records
echo For i = 0 To entCount - 1
echo     bgl = entLines^(i^).Item^("BGL"^)
echo     amount = entLines^(i^).Item^("AMOUNT"^)
echo     
echo     ' Get SL NO from business rules
echo     slNo = GetSlNo^(bgl, amount^)
echo     
echo     ' Get TAT from IOR data
echo     tat = ""
echo     If iorDict.Exists^(bgl^) Then
echo         tat = iorDict^(bgl^)^(0^)
echo     End If
echo     
echo     ' Calculate days passed and overdue
echo     Dim postDateStr, daysPassed, overdueVal
echo     postDateStr = entLines^(i^).Item^("POST_DATE"^)
echo     daysPassed = 0
echo     overdueVal = "NA"
echo     
echo     If postDateStr ^<^> "X" And postDateStr ^<^> "" Then
echo         ' Simple date calculation ^(assuming YYYY-MM-DD format^)
echo         On Error Resume Next
echo         postDate = CDate^(postDateStr^)
echo         If Err.Number = 0 Then
echo             daysPassed = DateDiff^("d", postDate, currentDate^)
echo             If IsNumeric^(tat^) And tat ^<^> "" Then
echo                 overdueVal = daysPassed - CInt^(tat^)
echo             End If
echo         End If
echo         On Error GoTo 0
echo     End If
echo     
echo     ' Write record
echo     outputFile.WriteLine entLines^(i^).Item^("BRCD"^) ^& "," ^& _
echo                         entLines^(i^).Item^("BRANCH"^) ^& "," ^& _
echo                         entLines^(i^).Item^("CIRCLE"^) ^& "," ^& _
echo                         entLines^(i^).Item^("NETWORK"^) ^& "," ^& _
echo                         entLines^(i^).Item^("MOD_NO"^) ^& "," ^& _
echo                         entLines^(i^).Item^("MODULE"^) ^& "," ^& _
echo                         entLines^(i^).Item^("REG_NO"^) ^& "," ^& _
echo                         entLines^(i^).Item^("BGL"^) ^& "," ^& _
echo                         entLines^(i^).Item^("BGLDESC"^) ^& "," ^& _
echo                         entLines^(i^).Item^("REF"^) ^& "," ^& _
echo                         entLines^(i^).Item^("AMOUNT"^) ^& "," ^& _
echo                         postDateStr ^& "," ^& _
echo                         entLines^(i^).Item^("DAYS"^) ^& "," ^& _
echo                         tat ^& "," ^& _
echo                         slNo ^& "," ^& _
echo                         overdueVal ^& "," ^& _
echo                         daysPassed
echo     
echo     processedCount = processedCount + 1
echo Next
echo.
echo outputFile.Close
echo.
echo WScript.Echo "  Successfully processed " ^& processedCount ^& " records"
echo WScript.Echo "  Output saved to: output\Merged_Output.csv"
) > csv_processor.vbs

REM [4/9] RUNNING VBSCRIPT PROCESSOR
echo.
echo [4/9] RUNNING VBSCRIPT PROCESSOR...

cscript //nologo csv_processor.vbs
set "vbs_result=!errorlevel!"

REM [5/9] CLEANING UP
echo.
echo [5/9] CLEANING UP...
del csv_processor.vbs

if !vbs_result! neq 0 (
    echo [ERROR] VBScript processing failed!
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
echo 2. Verify CSV files are not corrupted or locked
echo 3. Make sure you have write permissions
echo 4. Close any Excel files using the output
echo 5. Ensure VBScript is enabled (should be by default)
echo.
pause
exit /b 1

:normal_exit
echo Press any key to exit...
pause >nul
exit /b 0