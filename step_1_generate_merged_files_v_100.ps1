<#
.SYNOPSIS
    Merges data from BAL, ENT, and IOR CSV files with comprehensive error handling
    and debugging features.

.TUTORIAL
    This script demonstrates:
    1. Directory structure management
    2. File validation
    3. CSV data processing
    4. Data merging techniques
    5. Business rule implementation
    6. Debugging practices
    7. Error handling

.NOTES
    File Structure Requirements:
    - Current Directory: Contains this script and IOR*.cdcsv
    - \input: Contains BAL*.csv and ENT*.csv
    - \output: Auto-created for final results
#>

[CmdletBinding()]
Param(
    [Parameter(Mandatory=$false)]
    [string]$OutputFile = '.\output\Merged_Output.csv'
)

#region INITIAL SETUP
#-----------------------------------------------------------
# SECTION 1: Script Initialization
# Purpose: Set up directories and basic script environment
#-----------------------------------------------------------
Write-Host "`n[1/9] INITIALIZING SCRIPT..." -ForegroundColor Cyan

# Create output directory if missing
$outputDirectory = Split-Path -Path $OutputFile -Parent
if (-not (Test-Path -Path $outputDirectory)) {
    Write-Host "  Creating output directory: $outputDirectory" -ForegroundColor DarkGray
    New-Item -ItemType Directory -Path $outputDirectory | Out-Null
}

# Store today's date for date calculations
$todayDate = Get-Date
#endregion

#region FILE DISCOVERY
#-----------------------------------------------------------
# SECTION 2: File Discovery
# Purpose: Locate required input files
#-----------------------------------------------------------
Write-Host "`n[2/9] LOCATING FILES..." -ForegroundColor Yellow

# Input directory configuration
$inputDirectory = ".\input"

# Find required files
$balFile = Get-ChildItem -Path $inputDirectory -Filter "BAL*.csv" -ErrorAction SilentlyContinue | Select-Object -First 1
$entFile = Get-ChildItem -Path $inputDirectory -Filter "ENT*.csv" -ErrorAction SilentlyContinue | Select-Object -First 1
$iorFile = Get-ChildItem -Path . -Filter "IOR*.csv" -ErrorAction SilentlyContinue | Select-Object -First 1

# Validate file existence
$missingFiles = @()
if (-not $balFile) { $missingFiles += "BAL*.csv in input folder" }
if (-not $entFile) { $missingFiles += "ENT*.csv in input folder" }
if (-not $iorFile) { $missingFiles += "IOR*.csv in current folder" }

if ($missingFiles.Count -gt 0) {
    Write-Host "`n[ERROR] Missing required files:" -ForegroundColor Red
    $missingFiles | ForEach-Object { Write-Host "  - $_" -ForegroundColor Red }
    Write-Host "`nPress any key to exit..." -ForegroundColor DarkGray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

Write-Host "`n[FILE LOCATIONS]" -ForegroundColor Green
Write-Host "  BAL File: $($balFile.FullName)"
Write-Host "  ENT File: $($entFile.FullName)"
Write-Host "  IOR File: $($iorFile.FullName)"
Write-Host "  Output File: $OutputFile"
#endregion

#region DATA PROCESSING
#-----------------------------------------------------------
# SECTION 3-7: Core Data Processing
# Purpose: Process and combine data from different sources
#-----------------------------------------------------------
try {
    #region BAL DATA PROCESSING
    #-------------------------------------------------------
    # SECTION 3: BAL Data Handling
    # Purpose: Process BAL file and create placeholder columns
    #-------------------------------------------------------
    Write-Host "`n[3/9] PROCESSING BAL DATA..." -ForegroundColor Yellow
  #-- $balData=@()
    $balData = Import-Csv -Path $balFile.FullName | Where-Object {
        $_.MODULE -eq 'TRICHY' -and
        $_.BGL -in ('98556', '98585', '98641')
    }

    # Replace the existing $balData Select-Object block with:
#--$balData=@()    
$balData = $balData | Select-Object *,
    @{Name='REF'; Expression={'X'}},
    @{Name='AMOUNT'; Expression={$_.BALANCE}},
    @{Name='POST DATE'; Expression={'X'}},
    @{Name='DAYS'; Expression={'X'}}


    Write-Host "  BAL records after filtering: $($balData.Count)"
    #endregion

    #region ENT DATA PROCESSING
    #-------------------------------------------------------
    # SECTION 4: ENT Data Handling
    # Purpose: Process ENT file with TRICHY filter
    #-------------------------------------------------------
    Write-Host "`n[4/9] PROCESSING ENT DATA..." -ForegroundColor Yellow
   
    $entData = Import-Csv -Path $entFile.FullName | Where-Object {
        $_.MODULE -eq 'TRICHY'
    }

    Write-Host "  ENT records after filtering: $($entData.Count)"
    #endregion

    #region DATA COMBINATION
    #-------------------------------------------------------
    # SECTION 5: Data Combination
    # Purpose: Combine BAL and ENT data
    #-------------------------------------------------------
    Write-Host "`n[5/9] COMBINING DATA..." -ForegroundColor Yellow
    $combinedData =@()
    $combinedData = $balData + $entData

    Write-Host "  Total combined records: $($combinedData.Count)"
    #endregion

    #region IOR DATA PROCESSING
    #-------------------------------------------------------
    # SECTION 6: IOR Data Handling
    # Purpose: Process IOR file with header correction
    #-------------------------------------------------------
    Write-Host "`n[6/9] PROCESSING IOR DATA..." -ForegroundColor Yellow
   
    $iorData = Get-Content -Path $iorFile.FullName |
        Select-Object -Skip 1 |
        ConvertFrom-Csv -Header "SLNO_A","BGL","BGL_NAME","TAT","SLNO_B" |
        Select-Object BGL, TAT, @{Name='SL NO'; Expression={$_.SLNO_B}}

    Write-Host "  IOR records processed: $($iorData.Count)"
    #endregion

    #region INITIAL MERGE
# -------------------------------------------------------
# SECTION 7: Initial Data Merge
# Purpose: Create base dataset with SL NO and TAT
#         (Now also applying BGL-based rules for SL NO)
# -------------------------------------------------------
Write-Host "`n[7/9] PERFORMING INITIAL MERGE..." -ForegroundColor Yellow

$processedData = foreach ($record in $combinedData) {
    # Find matching IOR record
    $iorMatch = $iorData | Where-Object { $_.BGL -eq $record.BGL } | Select-Object -First 1

    # Default SL NO from IOR or 9
    $slNo = if ($iorMatch) { $iorMatch.'SL NO' } else { 9 }

    # TAT from IOR if available
    $tat = if ($iorMatch) { $iorMatch.TAT } else { $null }

    # -- BUSINESS RULE LOGIC MOVED HERE --
    # Parse the 'AMOUNT' to apply BGL-based rules
    $amount = 0
    if (-not [string]::IsNullOrWhiteSpace($record.AMOUNT)) {
        try {
            # Remove non-numeric characters, then convert to [double]
            $amount = [math]::Abs([double]($record.AMOUNT -replace '[^\d.]',''))
        }
        catch {
            Write-Host "[WARNING] Could not parse AMOUNT for record with BGL $($record.BGL)" -ForegroundColor Yellow
            # $amount remains 0
        }
    }

    # Use a switch statement to override SL NO based on BGL & amount rules
    switch ($record.BGL.Trim()) {
        '98533' {
            if ($amount -ge 1e7) { $slNo = 4 } else { $slNo = 5 }
        }
        '2399869' {
            if ($amount -ge 1e7) { $slNo = 4 } else { $slNo = 6 }
        }
        '98593' {
            if ($amount -ge 1e7) { $slNo = 4 } else { $slNo = 5 }
        }
        Default {
            # Keep the existing $slNo as assigned above
        }
    }

    # Now create our final merged object *with* the correct SL NO
    [PSCustomObject]@{
        BRCD         = $record.BRCD
        BRANCH       = $record.BRANCH
        CIRCLE       = $record.CIRCLE
        NETWORK      = $record.NETWORK
        'MOD NO'     = $record.'MOD NO'
        MODULE       = $record.MODULE
        'REG NO'     = $record.'REG NO'
        BGL          = $record.BGL
        BGLDESC      = $record.BGLDESC
        REF          = $record.REF
        AMOUNT       = $record.AMOUNT
        'POST DATE'  = $record.'POST DATE'
        DAYS         = $record.DAYS
        TAT          = $tat
        'SL NO'      = $slNo
        OVERDUE      = $null
        'DAYS PASSED'= $null
    }
}

# Debug: Show first record's properties
Write-Host "`n[DEBUG] First Record Structure:" -ForegroundColor Magenta
if ($processedData.Count -gt 0) {
    $processedData[0].PSObject.Properties | ForEach-Object {
        Write-Host "  $($_.Name.PadRight(15)) = $($_.Value)" -ForegroundColor DarkGray
    }
}
else {
    Write-Host "  No records found!" -ForegroundColor Red
}
#endregion


<#
    #region INITIAL MERGE
    #-------------------------------------------------------
    # SECTION 7: Initial Data Merge
    # Purpose: Create base dataset with SL NO and TAT
    #-------------------------------------------------------
    Write-Host "`n[7/9] PERFORMING INITIAL MERGE..." -ForegroundColor Yellow
   
    $processedData = foreach ($record in $combinedData) {
        $iorMatch = $iorData | Where-Object { $_.BGL -eq $record.BGL } | Select-Object -First 1
       
        # Create object with all required properties
        [PSCustomObject]@{
            BRCD        = $record.BRCD
            BRANCH      = $record.BRANCH
            CIRCLE      = $record.CIRCLE
            NETWORK     = $record.NETWORK
            'MOD NO'    = $record.'MOD NO'
            MODULE      = $record.MODULE
            'REG NO'    = $record.'REG NO'
            BGL         = $record.BGL
            BGLDESC     = $record.BGLDESC
            REF         = $record.REF
            AMOUNT      = $record.AMOUNT
            'POST DATE' = $record.'POST DATE'
            DAYS        = $record.DAYS
            TAT         = if ($iorMatch) { $iorMatch.TAT } else { $null }
            'SL NO'     = if ($iorMatch) { $iorMatch.'SL NO' } else { 9 }
            OVERDUE     = $null
            'DAYS PASSED' = $null
        }
    }

    # Debug: Show first record's properties
    Write-Host "`n[DEBUG] First Record Structure:" -ForegroundColor Magenta
    if ($processedData.Count -gt 0) {
        $processedData[0].PSObject.Properties | ForEach-Object {
            Write-Host "  $($_.Name.PadRight(15)) = $($_.Value)" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "  No records found!" -ForegroundColor Red
    }
    #endregion
#>

<#
#region BUSINESS RULES
#-------------------------------------------------------
# SECTION 8: Business Rule Application
#-------------------------------------------------------
Write-Host "`n[8/9] APPLYING BUSINESS RULES..." -ForegroundColor Yellow


# In the "BUSINESS RULES" section, replace the switch logic with:
$processedData = $processedData | ForEach-Object {
    # Debug: Log current record
    Write-Host "Processing BGL: '$($_.BGL)' | AMOUNT: '$($_.AMOUNT)'" -ForegroundColor Cyan

    # Skip invalid records
    if (-not $_ -or [string]::IsNullOrWhiteSpace($_.BGL)) {
        Write-Host "[WARNING] Skipping invalid record" -ForegroundColor Yellow
        return $_
    }

    try {
        # Parse amount safely
        $amount = [double]0
        if (-not [string]::IsNullOrWhiteSpace($_.AMOUNT)) {
            $amount = [double]($_.AMOUNT -replace '[^\d.]', '')
        }

        # Debug: Log parsed amount
        Write-Host "  Parsed AMOUNT: $amount" -ForegroundColor DarkCyan

        # Force-update SL NO based on BGL and amount
        switch ($_.BGL.Trim()) {
            '98533' {
                $newSlNo = if ($amount -ge 10000000) { 4 } else { 5 }
                $_ | Add-Member -MemberType NoteProperty -Name 'SL NO' -Value $newSlNo -Force
                Write-Host "  Updated BGL $($_.BGL) SL NO to $newSlNo" -ForegroundColor Green
            }
            '2399869' {
                $newSlNo = if ($amount -ge 10000000) { 4 } else { 6 }
                $_ | Add-Member -MemberType NoteProperty -Name 'SL NO' -Value $newSlNo -Force
                Write-Host "  Updated BGL $($_.BGL) SL NO to $newSlNo" -ForegroundColor Green
            }
            '98593' {
                $newSlNo = if ($amount -ge 10000000) { 4 } else { 5 }
                $_ | Add-Member -MemberType NoteProperty -Name 'SL NO' -Value $newSlNo -Force
                Write-Host "  Updated BGL $($_.BGL) SL NO to $newSlNo" -ForegroundColor Green
            }
            default {
                Write-Host "  No rule for BGL $($_.BGL). Keeping SL NO: $($_.'SL NO')" -ForegroundColor Gray
            }
        }
    }
    catch {
        Write-Host "[ERROR] Failed to process BGL $($_.BGL): $($_.Exception.Message)" -ForegroundColor Red
    }

    $_  # Return modified object
}

$processedData = $processedData | ForEach-Object {
    # Skip null records completely
    if ($_ -eq $null) {
        Write-Host "[WARNING] Skipping null record" -ForegroundColor Yellow
        return $null
    }

    # Force-create SL NO property if missing (with error handling)
    try {
        if (-not $_.PSObject.Properties['SL NO']) {
            Write-Host "[WARNING] Creating missing SL NO for record:" -ForegroundColor Yellow
            Write-Host "  BGL: $($_.BGL), AMOUNT: $($_.AMOUNT)" -ForegroundColor Yellow
            $_ | Add-Member -MemberType NoteProperty -Name 'SL NO' -Value 9 -Force
        }
    }
    catch {
        Write-Host "[CRITICAL ERROR] Failed to create SL NO property:" -ForegroundColor Red
        Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }

    # Skip records with empty BGL
    if ([string]::IsNullOrWhiteSpace($_.BGL)) {
        Write-Host "[WARNING] Skipping record with empty BGL:" -ForegroundColor Yellow
        Write-Host "  AMOUNT: $($_.AMOUNT), POST DATE: $($_.'POST DATE')" -ForegroundColor Yellow
        return $null
    }



    # Process valid records
    try {
        $amount = [double]0
        if (-not [string]::IsNullOrWhiteSpace($_.AMOUNT)) {
            $amount = [math]::Abs([double]($_.AMOUNT -replace '[^\d.]', ''))
        }

        Write-Host "Processing BGL: $($_.BGL) | AMOUNT: $($_.AMOUNT)" -ForegroundColor Cyan

        # In the "BUSINESS RULES" section, modify the switch block to:
        switch ($_.BGL.Trim()) {
            { $_ -in '98533', '2399869', '98593' } {
                try {
                    $amount = [math]::Abs([double]($_.AMOUNT -replace '[^\d.]', ''))
           
                    # Explicitly update SL NO property
                    $newSlNo = switch ($_) {
                        '98533'   { if ($amount -ge 1e7) { Write-Host "inside" 4 } else { Write-Host "outside" 5 } }
                        '2399869' { if ($amount -ge 1e7) { 4 } else { 6 } }
                        '98593'   { if ($amount -ge 1e7) { 4 } else { 5 } }
                    }
           
                    # Force-update the property
                    $_ | Add-Member -MemberType NoteProperty -Name 'SL NO' -Value $newSlNo -Force
                }
                catch {
                    Write-Host "[ERROR] Failed to process BGL $($_.BGL): $_" -ForegroundColor Red
                }
            }
            default {
                Write-Host "[INFO] No rule for BGL $($_.BGL). Keeping SL NO: $($_.'SL NO')" -ForegroundColor Gray
            }
        }

    }
    catch {
        Write-Host "[ERROR] Failed to process record:" -ForegroundColor Red
        Write-Host "  BGL: $($_.BGL)" -ForegroundColor Red
        Write-Host "  AMOUNT: $($_.AMOUNT)" -ForegroundColor Red
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }

    $_  # Return modified object
} | Where-Object { $_ -ne $null }  # Remove null records

Write-Host "  Valid records after business rules: $($processedData.Count)"
#endregion
#>

#region FINAL CALCULATIONS
#-------------------------------------------------------
# SECTION 9: Final Calculations
#-------------------------------------------------------
Write-Host "`n[9/9] PERFORMING FINAL CALCULATIONS..." -ForegroundColor Yellow

$finalData = $processedData | ForEach-Object {
    # Date calculations with error handling

    $postDate = $null
    if (-not [string]::IsNullOrWhiteSpace($_.'POST DATE')) {
        try {
            $postDate = [datetime]::ParseExact(
                $_.'POST DATE'.Trim(),
                'yyyy-MM-dd',
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
        catch {
            Write-Host "[WARNING] Invalid POST DATE format: $($_.'POST DATE')" -ForegroundColor Yellow
            Write-Host "  Using default date (today)" -ForegroundColor Yellow
            $postDate = $null
        }
    }


    # Days passed calculation
    $_.'DAYS PASSED' = if ($postDate) { ($todayDate - $postDate).Days } else { 0 }

<#    # TAT conversion with error handling
    $tatValue = 0
    if (-not [string]::IsNullOrWhiteSpace($_.TAT)) {
        try {
            $tatValue = [int]::Parse(
                $($_.TAT -replace '[^\d]', ''),  # Remove non-numeric characters
                [System.Globalization.NumberStyles]::Any,
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
        catch {
            Write-Host "[WARNING] Invalid TAT value: $($_.TAT)" -ForegroundColor Yellow
            Write-Host "  Using default TAT (0 days)" -ForegroundColor Yellow
            $tatValue = 0
        }
    }

    # Overdue calculation
    $_.OVERDUE = $_.'DAYS PASSED' - $tatValue #>

        # TAT conversion with error handling
    $tatValue = $null
    if (-not [string]::IsNullOrWhiteSpace($_.TAT)) {
        try {
            $tatValue = [int]::Parse(
                ($_.TAT -replace '[^\d]', ''),  # Remove non-numeric characters
                [System.Globalization.NumberStyles]::Any,
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
        catch {
            Write-Host "[WARNING] Invalid TAT value: $($_.TAT)" -ForegroundColor Yellow
            Write-Host "  Marking TAT as NA (non-numeric)" -ForegroundColor Yellow
            $tatValue = $null
        }
    }

    # Overdue calculation
    if ($tatValue -eq $null -Or $_.DAYS -eq "X") {
        $_.OVERDUE = 'NA'
    }
    else {
        $_.OVERDUE = $_.'DAYS' - $tatValue
    }


    $_  # Return modified object
}

#endregion

    #region DATA EXPORT
    #-------------------------------------------------------
    # FINAL STEP: Export Results
    # Purpose: Save final data to CSV
    #-------------------------------------------------------
    Write-Host "`n[EXPORTING RESULTS]" -ForegroundColor Green
    $finalData | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
    Write-Host "  Output saved to: $OutputFile" -ForegroundColor Cyan
    #endregion
}
catch {
    Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Error occurred in: $($_.InvocationInfo.ScriptName)" -ForegroundColor Red
    Write-Host "  Line: $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    pause
    exit
}

#region SCRIPT COMPLETION
#-----------------------------------------------------------
# Script Cleanup
# Purpose: Provide final feedback and keep window open
#-----------------------------------------------------------
Write-Host "`n[SCRIPT COMPLETED]" -ForegroundColor Green
Write-Host "  Output Directory: $outputDirectory"
Write-Host "  Total Records Processed: $($finalData.Count)"
Write-Host "`nPress any key to exit..." -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
#endregion
