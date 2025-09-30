' #########################################################
' VBScript: Convert Excel (.xlsx) to CSV
' #########################################################

' Step 1: Read input arguments
Dim objExcel, objWorkbook
Set objExcel = CreateObject("Excel.Application")
objExcel.DisplayAlerts = False

Dim inputFile, outputFile
inputFile = WScript.Arguments(0)  ' Path to .xlsx file
outputFile = WScript.Arguments(1) ' Path to .csv file

' Step 2: Open the Excel workbook
Set objWorkbook = objExcel.Workbooks.Open(inputFile)

' Step 3: Save the workbook as CSV
objWorkbook.SaveAs outputFile, 6  ' 6 = xlCSV format

' Step 4: Clean up
objWorkbook.Close False
objExcel.Quit
Set objWorkbook = Nothing
Set objExcel = Nothing
