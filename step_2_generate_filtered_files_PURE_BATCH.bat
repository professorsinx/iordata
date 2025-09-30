@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM                    STEP 2 - PURE BATCH + VBSCRIPT VERSION
REM ============================================================================
REM Purpose: Filters and sorts the merged CSV data from Step 1 into specific files
REM Input: output\Merged_Output.csv (from Step 1)
REM Output: Multiple filtered CSV files in output folder
REM Requirements: Windows with VBScript support (no external dependencies)
REM ============================================================================

echo.
echo ========================================
echo   STEP 2: CSV FILTER AND SORT
echo ========================================
echo Using VBScript for advanced CSV filtering
echo.

REM [1/3] INITIALIZING AND VALIDATING
echo [1/3] INITIALIZING AND VALIDATING...

REM Check if input file from Step 1 exists
if not exist "output\Merged_Output.csv" (
    echo [ERROR] Input file not found: output\Merged_Output.csv
    echo Please run Step 1 first to generate the merged data.
    goto :error_exit
)

REM Create output directory if it doesn't exist
if not exist "output" (
    echo   Creating output directory...
    mkdir output
)

echo   Input file found: output\Merged_Output.csv
echo   Ready to process filtered files...

REM [2/3] CREATING VBSCRIPT FILTER PROCESSOR
echo.
echo [2/3] CREATING VBSCRIPT FILTER PROCESSOR...

REM Create comprehensive VBScript for filtering and sorting
(
echo Option Explicit
echo.
echo Dim fso, inputFile, outputPath
echo Dim allData^(^), headers, recordCount
echo Dim i, j, line, fields, record
echo.
echo Set fso = CreateObject^("Scripting.FileSystemObject"^)
echo outputPath = "output\"
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
echo ' Function to write CSV line with proper escaping
echo Function WriteCSVLine^(fields^)
echo     Dim result, i, field
echo     result = ""
echo     For i = 0 To UBound^(fields^)
echo         field = CStr^(fields^(i^)^)
echo         If InStr^(field, ","^) ^> 0 Or InStr^(field, """"^) ^> 0 Then
echo             field = """" ^& Replace^(field, """", """"""^) ^& """"
echo         End If
echo         If i ^> 0 Then result = result ^& ","
echo         result = result ^& field
echo     Next
echo     WriteCSVLine = result
echo End Function
echo.
echo ' Function to sort array of records by specified columns
echo Sub SortRecords^(records, sortColumns^)
echo     Dim i, j, temp, swap
echo     Dim col1Val, col2Val
echo     
echo     ' Simple bubble sort ^(adequate for typical data sizes^)
echo     For i = 0 To UBound^(records^) - 1
echo         For j = i + 1 To UBound^(records^)
echo             swap = False
echo             
echo             ' Compare based on sort columns
echo             For Each col In sortColumns
echo                 col1Val = records^(i^).Item^(col^)
echo                 col2Val = records^(j^).Item^(col^)
echo                 
echo                 If col1Val ^> col2Val Then
echo                     swap = True
echo                     Exit For
echo                 ElseIf col1Val ^< col2Val Then
echo                     Exit For
echo                 End If
echo             Next
echo             
echo             If swap Then
echo                 Set temp = records^(i^)
echo                 Set records^(i^) = records^(j^)
echo                 Set records^(j^) = temp
echo             End If
echo         Next
echo     Next
echo End Sub
echo.
echo ' Function to export filtered data
echo Sub ExportFilteredData^(filteredRecords, fileName, description^)
echo     Dim outputFile, i
echo     Dim outputFields
echo     
echo     WScript.Echo "  Generating: " ^& fileName ^& " ^(" ^& description ^& "^)"
echo     
echo     Set outputFile = fso.CreateTextFile^(outputPath ^& fileName, True^)
echo     
echo     ' Write header
echo     outputFile.WriteLine "BRCD,BRANCH,REG NO,BGL,BGLDESC,REF,AMOUNT,POST DATE,DAYS,TAT,SL NO,OVERDUE,DAYS PASSED"
echo     
echo     ' Write filtered records
echo     For i = 0 To UBound^(filteredRecords^)
echo         outputFields = Array^( _
echo             filteredRecords^(i^).Item^("BRCD"^), _
echo             filteredRecords^(i^).Item^("BRANCH"^), _
echo             filteredRecords^(i^).Item^("REG_NO"^), _
echo             filteredRecords^(i^).Item^("BGL"^), _
echo             filteredRecords^(i^).Item^("BGLDESC"^), _
echo             filteredRecords^(i^).Item^("REF"^), _
echo             filteredRecords^(i^).Item^("AMOUNT"^), _
echo             filteredRecords^(i^).Item^("POST_DATE"^), _
echo             filteredRecords^(i^).Item^("DAYS"^), _
echo             filteredRecords^(i^).Item^("TAT"^), _
echo             filteredRecords^(i^).Item^("SL_NO"^), _
echo             filteredRecords^(i^).Item^("OVERDUE"^), _
echo             filteredRecords^(i^).Item^("DAYS_PASSED"^) _
echo         ^)
echo         outputFile.WriteLine WriteCSVLine^(outputFields^)
echo     Next
echo     
echo     outputFile.Close
echo     WScript.Echo "    Records exported: " ^& ^(UBound^(filteredRecords^) + 1^)
echo End Sub
echo.
echo ' Load input data
echo WScript.Echo "Loading merged data from Step 1..."
echo.
echo Set inputFile = fso.OpenTextFile^("output\Merged_Output.csv", 1^)
echo headers = inputFile.ReadLine  ' Read and store header
echo.
echo ' Load all records into memory
echo recordCount = 0
echo ReDim allData^(1000^)  ' Initial size
echo.
echo Do While Not inputFile.AtEndOfStream
echo     line = inputFile.ReadLine
echo     If Len^(Trim^(line^)^) ^> 0 Then
echo         fields = ParseCSVLine^(line^)
echo         If UBound^(fields^) ^>= 12 Then  ' Ensure sufficient fields
echo             Set record = CreateObject^("Scripting.Dictionary"^)
echo             record.Add "BRCD", fields^(0^)
echo             record.Add "BRANCH", fields^(1^)
echo             record.Add "CIRCLE", fields^(2^)
echo             record.Add "NETWORK", fields^(3^)
echo             record.Add "MOD_NO", fields^(4^)
echo             record.Add "MODULE", fields^(5^)
echo             record.Add "REG_NO", fields^(6^)
echo             record.Add "BGL", fields^(7^)
echo             record.Add "BGLDESC", fields^(8^)
echo             record.Add "REF", fields^(9^)
echo             record.Add "AMOUNT", fields^(10^)
echo             record.Add "POST_DATE", fields^(11^)
echo             record.Add "DAYS", fields^(12^)
echo             record.Add "TAT", fields^(13^)
echo             record.Add "SL_NO", fields^(14^)
echo             record.Add "OVERDUE", fields^(15^)
echo             record.Add "DAYS_PASSED", fields^(16^)
echo             
echo             Set allData^(recordCount^) = record
echo             recordCount = recordCount + 1
echo             If recordCount Mod 100 = 0 Then ReDim Preserve allData^(recordCount + 100^)
echo         End If
echo     End If
echo Loop
echo inputFile.Close
echo ReDim Preserve allData^(recordCount - 1^)
echo.
echo WScript.Echo "Loaded " ^& recordCount ^& " records for filtering"
echo.
echo ' Filter 1: SL NO = 9
echo WScript.Echo "Processing Filter 1: SL NO = 9"
echo Dim filter1Records^(^), filter1Count
echo filter1Count = 0
echo ReDim filter1Records^(recordCount^)
echo.
echo For i = 0 To recordCount - 1
echo     If Trim^(allData^(i^).Item^("SL_NO"^)^) = "9" Then
echo         Set filter1Records^(filter1Count^) = allData^(i^)
echo         filter1Count = filter1Count + 1
echo     End If
echo Next
echo ReDim Preserve filter1Records^(filter1Count - 1^)
echo.
echo If filter1Count ^> 0 Then
echo     SortRecords filter1Records, Array^("SL_NO", "REG_NO", "BRCD"^)
echo     ExportFilteredData filter1Records, "Filtered_SL_NO_9.csv", "SL NO = 9"
echo End If
echo.
echo ' Filter 2: SL NO in ^(1, 2^) AND DAYS PASSED ^> 20 AND TAT = 45
echo WScript.Echo "Processing Filter 2: SL NO in ^(1, 2^), DAYS PASSED ^> 20, TAT = 45"
echo Dim filter2Records^(^), filter2Count
echo filter2Count = 0
echo ReDim filter2Records^(recordCount^)
echo.
echo For i = 0 To recordCount - 1
echo     Dim slNo, daysPassed, tat
echo     slNo = Trim^(allData^(i^).Item^("SL_NO"^)^)
echo     daysPassed = Trim^(allData^(i^).Item^("DAYS_PASSED"^)^)
echo     tat = Trim^(allData^(i^).Item^("TAT"^)^)
echo     
echo     If ^(slNo = "1" Or slNo = "2"^) And IsNumeric^(daysPassed^) And IsNumeric^(tat^) Then
echo         If CInt^(daysPassed^) ^> 20 And CInt^(tat^) = 45 Then
echo             Set filter2Records^(filter2Count^) = allData^(i^)
echo             filter2Count = filter2Count + 1
echo         End If
echo     End If
echo Next
echo ReDim Preserve filter2Records^(filter2Count - 1^)
echo.
echo If filter2Count ^> 0 Then
echo     SortRecords filter2Records, Array^("SL_NO", "REG_NO", "BRCD"^)
echo     ExportFilteredData filter2Records, "Filtered_SL_NO_1_2_DaysPassed_GT_20.csv", "SL NO in ^(1, 2^), DAYS PASSED ^> 20, TAT = 45"
echo End If
echo.
echo ' Filter 3: TAT != 45 AND OVERDUE in ^(-1, -2, -3^)
echo WScript.Echo "Processing Filter 3: TAT != 45, OVERDUE in ^(-1, -2, -3^)"
echo Dim filter3Records^(^), filter3Count
echo filter3Count = 0
echo ReDim filter3Records^(recordCount^)
echo.
echo For i = 0 To recordCount - 1
echo     tat = Trim^(allData^(i^).Item^("TAT"^)^)
echo     Dim overdue
echo     overdue = Trim^(allData^(i^).Item^("OVERDUE"^)^)
echo     
echo     If IsNumeric^(tat^) And IsNumeric^(overdue^) Then
echo         If CInt^(tat^) ^<^> 45 And ^(CInt^(overdue^) = -1 Or CInt^(overdue^) = -2 Or CInt^(overdue^) = -3^) Then
echo             Set filter3Records^(filter3Count^) = allData^(i^)
echo             filter3Count = filter3Count + 1
echo         End If
echo     End If
echo Next
echo ReDim Preserve filter3Records^(filter3Count - 1^)
echo.
echo If filter3Count ^> 0 Then
echo     SortRecords filter3Records, Array^("SL_NO", "REG_NO", "BRCD"^)
echo     ExportFilteredData filter3Records, "Filtered_TAT_NE_45_Overdue_NEG_123.csv", "TAT != 45, OVERDUE in ^(-1, -2, -3^)"
echo End If
echo.
echo ' Filter 4: BGL = 4599635
echo WScript.Echo "Processing Filter 4: BGL = 4599635"
echo Dim filter4Records^(^), filter4Count
echo filter4Count = 0
echo ReDim filter4Records^(recordCount^)
echo.
echo For i = 0 To recordCount - 1
echo     If Trim^(allData^(i^).Item^("BGL"^)^) = "4599635" Then
echo         Set filter4Records^(filter4Count^) = allData^(i^)
echo         filter4Count = filter4Count + 1
echo     End If
echo Next
echo ReDim Preserve filter4Records^(filter4Count - 1^)
echo.
echo If filter4Count ^> 0 Then
echo     SortRecords filter4Records, Array^("SL_NO", "REG_NO", "BRCD"^)
echo     ExportFilteredData filter4Records, "Filtered_BGL_4599635.csv", "BGL = 4599635"
echo End If
echo.
echo ' Filter 5: BGL = 4597998
echo WScript.Echo "Processing Filter 5: BGL = 4597998"
echo Dim filter5Records^(^), filter5Count
echo filter5Count = 0
echo ReDim filter5Records^(recordCount^)
echo.
echo For i = 0 To recordCount - 1
echo     If Trim^(allData^(i^).Item^("BGL"^)^) = "4597998" Then
echo         Set filter5Records^(filter5Count^) = allData^(i^)
echo         filter5Count = filter5Count + 1
echo     End If
echo Next
echo ReDim Preserve filter5Records^(filter5Count - 1^)
echo.
echo If filter5Count ^> 0 Then
echo     SortRecords filter5Records, Array^("SL_NO", "REG_NO", "BRCD"^)
echo     ExportFilteredData filter5Records, "Filtered_BGL_4597998.csv", "BGL = 4597998"
echo End If
echo.
echo ' Filter 6: BGL = 4897932
echo WScript.Echo "Processing Filter 6: BGL = 4897932"
echo Dim filter6Records^(^), filter6Count
echo filter6Count = 0
echo ReDim filter6Records^(recordCount^)
echo.
echo For i = 0 To recordCount - 1
echo     If Trim^(allData^(i^).Item^("BGL"^)^) = "4897932" Then
echo         Set filter6Records^(filter6Count^) = allData^(i^)
echo         filter6Count = filter6Count + 1
echo     End If
echo Next
echo ReDim Preserve filter6Records^(filter6Count - 1^)
echo.
echo If filter6Count ^> 0 Then
echo     SortRecords filter6Records, Array^("SL_NO", "REG_NO", "BRCD"^)
echo     ExportFilteredData filter6Records, "Filtered_BGL_4897932.csv", "BGL = 4897932"
echo End If
echo.
echo ' Filter 7: SL NO != 9 AND OVERDUE ^>= 0
echo WScript.Echo "Processing Filter 7: SL NO != 9 AND OVERDUE ^>= 0"
echo Dim filter7Records^(^), filter7Count
echo filter7Count = 0
echo ReDim filter7Records^(recordCount^)
echo.
echo For i = 0 To recordCount - 1
echo     slNo = Trim^(allData^(i^).Item^("SL_NO"^)^)
echo     overdue = Trim^(allData^(i^).Item^("OVERDUE"^)^)
echo     
echo     If slNo ^<^> "9" And IsNumeric^(overdue^) Then
echo         If CInt^(overdue^) ^>= 0 Then
echo             Set filter7Records^(filter7Count^) = allData^(i^)
echo             filter7Count = filter7Count + 1
echo         End If
echo     End If
echo Next
echo ReDim Preserve filter7Records^(filter7Count - 1^)
echo.
echo If filter7Count ^> 0 Then
echo     SortRecords filter7Records, Array^("SL_NO", "REG_NO", "BRCD"^)
echo     ExportFilteredData filter7Records, "Filtered_SL_NO_NE_9_Overdue_GE_0.csv", "SL NO != 9 AND OVERDUE ^>= 0"
echo End If
echo.
echo WScript.Echo "All filtering operations completed successfully!"
) > step2_filter_processor.vbs

REM [3/3] RUNNING FILTER PROCESSOR
echo.
echo [3/3] RUNNING FILTER PROCESSOR...

cscript //nologo step2_filter_processor.vbs
set "vbs_result=!errorlevel!"

REM CLEANUP
echo.
echo CLEANING UP...
del step2_filter_processor.vbs

if !vbs_result! neq 0 (
    echo [ERROR] VBScript filtering failed!
    goto :error_exit
)

REM VERIFICATION AND SUMMARY
echo.
echo ========================================
echo   STEP 2 COMPLETED SUCCESSFULLY!
echo ========================================
echo.
echo Generated filtered files in output folder:
echo   1. Filtered_SL_NO_9.csv
echo   2. Filtered_SL_NO_1_2_DaysPassed_GT_20.csv
echo   3. Filtered_TAT_NE_45_Overdue_NEG_123.csv
echo   4. Filtered_BGL_4599635.csv
echo   5. Filtered_BGL_4597998.csv
echo   6. Filtered_BGL_4897932.csv
echo   7. Filtered_SL_NO_NE_9_Overdue_GE_0.csv
echo.

REM Count files in output directory
set "file_count=0"
for %%f in (output\Filtered_*.csv) do (
    set /a "file_count+=1"
)

echo Total filtered files created: !file_count!
echo Completion time: %time%
echo ========================================

goto :normal_exit

:error_exit
echo.
echo ========================================
echo   STEP 2 FAILED!
echo ========================================
echo.
echo TROUBLESHOOTING TIPS:
echo 1. Ensure Step 1 was completed successfully
echo 2. Check that output\Merged_Output.csv exists
echo 3. Verify you have write permissions to output folder
echo 4. Make sure CSV file is not open in Excel
echo 5. Ensure VBScript is enabled on your system
echo.
pause
exit /b 1

:normal_exit
echo.
echo Press any key to exit...
pause >nul
exit /b 0