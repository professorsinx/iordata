# Step 1: Load Merged Data and Remove Unnecessary Columns
Write-Host "[INFO] Step 1: Removing unnecessary columns from the merged file..."

# Import the merged file
$mergedFile = '.\output\Merged_Output.csv'  # Adjust this path if needed
$mergedData = Import-Csv -Path $mergedFile

# Select only the necessary columns
$filteredData = $mergedData | Select-Object `
    BRCD,
    BRANCH,
    'REG NO',
    BGL,
    BGLDESC,
    REF,
    AMOUNT,
    'POST DATE',
    DAYS,
    TAT,
    'SL NO',
    OVERDUE,
    'DAYS PASSED'

# Save the filtered data to a new file
$filteredFile = '.\output\Filtered_Merged_Output.csv'
$filteredData | Export-Csv -Path $filteredFile -NoTypeInformation -Encoding UTF8

Write-Host "[INFO] Unnecessary columns removed. Saved filtered data to $filteredFile."

#----------------------------------------------------------------------------------------------------

# Step 2: Generating Filtered and Sorted Files
Write-Host "[INFO] Step 2: Generating filtered and sorted files..."

# Load the filtered data from Step 1
$filteredFile = '.\output\Filtered_Merged_Output.csv'  # Adjust path if necessary
$filteredData = Import-Csv -Path $filteredFile

# Function to export filtered and sorted data
function Export-FilteredData {
    param (
        [array]$Data,                      # The data to filter and export
        [string]$FilterDescription,        # Description of the filter applied
        [string]$OutputFileName,           # Name of the output CSV file
        [string[]]$SortColumns             # Columns to sort the data
    )
    Write-Host "[INFO] Generating file: $OutputFileName ($FilterDescription)"
    # Sort the data and export it as a CSV file
    $Data | Sort-Object -Property $SortColumns | Export-Csv -Path $OutputFileName -NoTypeInformation -Encoding UTF8
}


# File 1: Filtered where SL NO = 9
$file1Data = $filteredData | Where-Object { $_.'SL NO' -eq '9' }
Export-FilteredData -Data $file1Data `
    -FilterDescription "SL NO = 9" `
    -OutputFileName '.\output\Filtered_SL_NO_9.csv' `
    -SortColumns @('SL NO', 'REG NO', 'BRCD')

# File 2: Filtered where SL NO in (1, 2) AND DAYS PASSED > 20 AND TAT = 45 AND BGL != 98581
$file2Data = $filteredData | Where-Object {
    ($_. 'SL NO' -eq '1' -or $_.'SL NO' -eq '2') -and
    ([int]$_. 'DAYS PASSED' -gt 20) -and
    ($_.TAT -eq '45') -and
    ($_.BGL -ne '98581')
}
Export-FilteredData -Data $file2Data `
   # -FilterDescription "SL NO in (1, 2), DAYS PASSED > 20, TAT = 45, BGL != 98581" `
    -FilterDescription "SL NO in (1, 2), DAYS PASSED > 20, TAT = 45" `
    -OutputFileName '.\output\Filtered_SL_NO_1_2_DaysPassed_GT_20.csv' `
    -SortColumns @('SL NO', 'REG NO', 'BRCD')

# File 3: Filtered where TAT != 45 AND OVERDUE in (-1, -2, -3)
$file3Data = $filteredData | Where-Object {
    ($_.TAT -ne '45') -and
    ($_.OVERDUE -in @('-1', '-2', '-3'))
}
Export-FilteredData -Data $file3Data `
    -FilterDescription "TAT != 45, OVERDUE in (-1, -2, -3)" `
    -OutputFileName '.\output\Filtered_TAT_NE_45_Overdue_NEG_123.csv' `
    -SortColumns @('SL NO', 'REG NO', 'BRCD')

# File 4: Filtered where BGL = 4599635
$file4Data = $filteredData | Where-Object { $_.BGL -eq '4599635' }
Export-FilteredData -Data $file4Data `
    -FilterDescription "BGL = 4599635" `
    -OutputFileName '.\output\Filtered_BGL_4599635.csv' `
    -SortColumns @('SL NO', 'REG NO', 'BRCD')

# File 5: Filtered where BGL = 4597998
$file5Data = $filteredData | Where-Object { $_.BGL -eq '4597998' }
Export-FilteredData -Data $file5Data `
    -FilterDescription "BGL = 4597998" `
    -OutputFileName '.\output\Filtered_BGL_4597998.csv' `
    -SortColumns @('SL NO', 'REG NO', 'BRCD')

# File 6: Filtered where BGL = 4897932
$file6Data = $filteredData | Where-Object { $_.BGL -eq '4897932' }
Export-FilteredData -Data $file6Data `
    -FilterDescription "BGL = 4897932" `
    -OutputFileName '.\output\Filtered_BGL_4897932.csv' `
    -SortColumns @('SL NO', 'REG NO', 'BRCD')

#----------------------------------------------------------------------------------------------------

# NEW FILTER AND EXPORT FOR FILE 7
# File 7: Filtered where SL NO is not equal to 9 and OVERDUE > 0

# Step 1: Apply the filter
# Use Where-Object to filter records where SL NO is not '9' and OVERDUE is greater than 0
$file7Data = $filteredData | Where-Object {
    ($_. 'SL NO' -ne '9') -and ([int]$_.OVERDUE -ge 0)
}

# Step 2: Provide filter description and file name
# Description for the filter applied
$file7Description = "SL NO != 9 and OVERDUE >= 0"
# Name of the output CSV file
$file7OutputFileName = '.\output\Filtered_SL_NO_NE_9_Overdue_GE_0.csv'

# Step 3: Export the filtered data
Export-FilteredData -Data $file7Data `
    -FilterDescription $file7Description `
    -OutputFileName $file7OutputFileName `
    -SortColumns @('SL NO', 'REG NO', 'BRCD')

Write-Host "[INFO] Generated File 7: SL NO != 9 and OVERDUE >= 0, saved as $file7OutputFileName."

#----------------------------------------------------------------------------------------------------


Write-Host "[INFO] All filtered and sorted files generated successfully."
