Dim objExcel, objWorkbook, objFSO, objFolder, objFile, objSheet
Dim outputDir, csvFile, sheetName, lastRow
Dim objArgs

Const MAX_SHEET_NAME_LENGTH = 31

' Get the output directory from arguments
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    WScript.Echo "Error: No output directory provided."
    WScript.Quit
End If
outputDir = objArgs(0)

' Create Excel objects
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False
Set objWorkbook = objExcel.Workbooks.Add

' Prepare for file lookups
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(outputDir)

' Keep track of used sheet names to avoid duplicates
Dim usedSheetNames
Set usedSheetNames = CreateObject("Scripting.Dictionary")

Dim firstSheetUsed
firstSheetUsed = False

'-------------------------------------------'
' Helper function to clean/truncate the name'
'-------------------------------------------'
Function CleanSheetName(originalName)
    Dim tempName, baseName, counter
   
    ' Remove invalid characters
    tempName = Replace(originalName, "[", "_")
    tempName = Replace(tempName, "]", "_")
    tempName = Replace(tempName, ":", "_")
    tempName = Replace(tempName, "\", "_")
    tempName = Replace(tempName, "/", "_")
    tempName = Replace(tempName, "?", "_")
    tempName = Replace(tempName, "*", "_")

    ' Truncate to max length
    If Len(tempName) > MAX_SHEET_NAME_LENGTH Then
        tempName = Left(tempName, MAX_SHEET_NAME_LENGTH)
    End If

    ' If the name ended up empty, assign a default
    If tempName = "" Then
        tempName = "Sheet"
    End If

    ' Ensure uniqueness by appending (2), (3), etc. if necessary
    baseName = tempName
    counter = 2
    While usedSheetNames.Exists(tempName)
        tempName = baseName & "(" & counter & ")"
        ' If that suffix again makes it exceed 31 chars, we trim it
        If Len(tempName) > MAX_SHEET_NAME_LENGTH Then
            tempName = Left(baseName, MAX_SHEET_NAME_LENGTH - Len("(" & counter & ")")) & "(" & counter & ")"
        End If
        counter = counter + 1
    Wend
   
    usedSheetNames.Add tempName, True
    CleanSheetName = tempName
End Function

'------------------------------------'
' Main loop over Filtered*.csv files '
'------------------------------------'
For Each objFile In objFolder.Files
    If LCase(Left(objFile.Name, 8)) = "filtered" And LCase(Right(objFile.Name, 4)) = ".csv" Then
       
        csvFile = objFile.Path
        sheetName = objFSO.GetBaseName(objFile.Name) ' e.g., "FilteredXYZ"
       
        ' Get a cleaned/truncated/unique sheet name
        sheetName = CleanSheetName(sheetName)

        ' Reuse the first default sheet or add a new one
        If Not firstSheetUsed Then
            firstSheetUsed = True
            objWorkbook.Sheets(1).Name = sheetName
            Set objSheet = objWorkbook.Sheets(1)
        Else
            Set objSheet = objWorkbook.Sheets.Add
            objSheet.Name = sheetName
        End If

        ' Put a note in cell A1 indicating the source file
  '      objSheet.Cells(1, 1).Value = "Imported: " & objFile.Name

        ' Import CSV data starting in A2 (create new QueryTable object)
        Dim newQTable
        Set newQTable = objSheet.QueryTables.Add("TEXT;" & csvFile, objSheet.Range("A1"))
        newQTable.TextFileCommaDelimiter = True
        newQTable.Refresh

    End If
Next

' Save workbook in the output directory as MergedExcel.xlsx
Dim excelFilePath
excelFilePath = outputDir & "\MergedExcel.xlsx"
objWorkbook.SaveAs excelFilePath, 51 ' 51 = xlOpenXMLWorkbook (.xlsx)
objWorkbook.Close False
objExcel.Quit

' 2) Delete the CSV files (still have objFolder, objFSO in scope here)
Dim f
For Each f In objFolder.Files
    If LCase(Left(f.Name, 8)) = "filtered" And LCase(Right(f.Name, 4)) = ".csv" Then
        f.Delete
    End If
Next

' 3) Now do the cleanup
Set objSheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
Set objFSO = Nothing
Set objFolder = Nothing
Set objFile = Nothing
Set usedSheetNames = Nothing

WScript.Echo "Excel file created successfully: " & excelFilePath
