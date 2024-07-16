##################################################################
## Copyright 2024  Stephen Fearns
##################################################################

Clear-Host
$Started = Get-Date

# Constants
$Left = 2; $Center = 3; $Right = 4
$Top = 1;  $Middle = 2; $Bottom = 3

# Microsoft.Office.Interop.Excel.XlBorderWeight
$xlHairline = 1; $xlThin = 2; $xlThick = 4; $xlMedium = -4138

# Microsoft.Office.Interop.Excel.XlBordersIndex
$xlDiagonalDown   = 5;  $xlDiagonalUp       = 6
$xlEdgeLeft       = 7;  $xlEdgeTop          = 8; $xlEdgeBottom = 9; $xlEdgeRight = 10
$xlInsideVertical = 11; $xlInsideHorizontal = 12

# Microsoft.Office.Interop.Excel.XlLineStyle
$xlContinuous    = 1;     $xlDashDot   = 4;     $xlDashDotDot = 5;     $xlSlantDashDot = 13
$xlLineStyleNone = -4142; $xlDouble    = -4119; $xlDot        = -4118; $xlDash         = -4115
$xlShiftToRight  = -4161; $xlShiftDown = -4121

# Color index
$xlColorIndexBlue = 5 # depends on default palette

# Spreadsheet(s) details
$PasswordProtectionRequired = $false
$SpreadsheetPassword = "PasswordGoesHere"  # This password is only to protect parts of the spreadsheet from changes which is why its been entered in clear-text

# Full path
$CurrentPath = "C:\Temp"
# Current folder
$CurrentPath = "."

$InputFolder  = $CurrentPath + "\Input"
$OutputFolder = $CurrentPath + "\Output"
$ScreenOutput = $OutputFolder + "\_Screen output.txt"
$ProcessingNotesFile = $OutputFolder + "\_Processing notes.txt"

Write-Output "This script processes old-style Risk Assessments`n" | Tee-Object -FilePath $ScreenOutput
Write-Output "Started:       $($Started.DateTime)`n" | Tee-Object -Append -FilePath $ScreenOutput

"This file contains notes about the source files as they were being processed`n" > $ProcessingNotesFile
"All these notes have been automatically created by our script" >> $ProcessingNotesFile
"----------------------------------------------------------------------------" >> $ProcessingNotesFile

Write-Output "Source Folder: $($InputFolder)" | Tee-Object -Append -FilePath $ScreenOutput
Write-Output "Output Folder: $($OutputFolder)`n" | Tee-Object -Append -FilePath $ScreenOutput

$TemplateSpreadsheet = $CurrentPath + "\Risk Assessment Template v1.xlsx"
Write-Output "Template:      $($TemplateSpreadsheet)" | Tee-Object -Append -FilePath $ScreenOutput

# Get the list of files in the $InputFolder
[array]$InputFileList = Get-ChildItem -Path $InputFolder
$InputFileCount = $InputFileList.Count
Write-Output "File Count:    $($InputFileCount)`n" | Tee-Object -Append -FilePath $ScreenOutput

# Used by the progress bar
$FilesProgressed = 0

ForEach ($File in $InputFileList) {
	# Excel pointer
	$InputExcel  = New-Object -ComObject Excel.Application
	$OutputExcel = New-Object -ComObject Excel.Application

	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($InputExcel.Version)\Excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null
	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($InputExcel.Version)\Excel\Security" -Name VBAWarnings -Value 1 -Force | Out-Null

    # Update the progress bar
    $FilesProgressed++
    Write-Progress -PercentComplete ($FilesProgressed/$InputFileCount*100) -Activity "Processing Old-Style Risk Assessments" -Status "File $($FilesProgressed) of $($InputFileCount)" | Tee-Object -Append -FilePath $ScreenOutput

    Write-Output "`nProcessing:    $($File.Name)" | Tee-Object -Append -FilePath $ScreenOutput

    $InputFileName  = $InputFolder + "\" + $File.Name
    Write-Output "  Input:       $($InputFileName)" | Tee-Object -Append -FilePath $ScreenOutput
    $OutputFileName = $OutputFolder + "\" + ($File.Name -Replace ".xlsx","__unsigned.xlsx")
    Write-Output "  Output:      $($OutputFileName)" | Tee-Object -Append -FilePath $ScreenOutput

    # Change $true to $false if Excel should not be visible
    $InputExcel.Visible  = $false
    $OutputExcel.Visible = $false
    $InputExcel.ScreenUpdating  = $true
    $OutputExcel.ScreenUpdating = $true

    # Change $true to $false if you don't want to see any alerts from Excel
    $InputExcel.DisplayAlerts  = $true
    $OutputExcel.DisplayAlerts = $true

    # Open the Spreadsheets now Excel has loaded
    # Open parameters : Filename, UpdateLinks, ReadOnly, Format, Password, ....)
    $InputWorkBook = $InputExcel.Workbooks.Open($InputFileName, $false, $false, [Type]::Missing, $SpreadsheetPassword)
    $OutputWorkBook = $OutputExcel.Workbooks.Open($TemplateSpreadsheet)

    # Select the sheet we are working with as there is more than 1
    $InputSheet  = $InputWorkBook.WorkSheets.Item(1)
    $OutputSheet = $OutputWorkBook.WorkSheets.Item(1)

	if ($PasswordProtectionRequired) {
		# Unlock the template file
		$InputSheet.Unprotect($SpreadsheetPassword)
	}

    # Remove the files from if the name matches?
    if (Test-Path $OutputFileName) {
        Remove-Item $OutputFileName -Force
        Write-Output "  Removed old: $($OutputFileName)" | Tee-Object -Append -FilePath $ScreenOutput
    }

    # Populate the Output Spreadsheet with data from the Input Spreadsheet
    # xx.xx.Item(Row, Column)

    # Header section of the Spreadsheet
    Write-Output "    Header section being copied" | Tee-Object -Append -FilePath $ScreenOutput
    $OutputSheet.Cells.Item(1,8)   = $InputSheet.Cells.Item(1,6)  # WORK AREA
    $OutputSheet.Cells.Item(1,33)  = $InputSheet.Cells.Item(1,29) # DOC NO.
    $OutputSheet.Cells.Item(3,8)   = $InputSheet.Cells.Item(3,6)  # ACTIVITY/TASK
    $OutputSheet.Cells.Item(3,33)  = $InputSheet.Cells.Item(3,29) # DATE
    $OutputSheet.Cells.Item(6,10)  = $InputSheet.Cells.Item(7,8)  # ASSOCIATED DOCUMENTATION
    Write-Output "      Work Area:     $($InputSheet.Cells.Item(1,6).Text)"  | Tee-Object -Append -FilePath $ScreenOutput
    Write-Output "      Doc No:        $($InputSheet.Cells.Item(1,29).Text)" | Tee-Object -Append -FilePath $ScreenOutput
    Write-Output "      Activity/Task: $($InputSheet.Cells.Item(3,6).Text)"  | Tee-Object -Append -FilePath $ScreenOutput
    Write-Output "      Date:          $($InputSheet.Cells.Item(3,29).Text)" | Tee-Object -Append -FilePath $ScreenOutput
    Write-Output "      Ass Docs:      $($InputSheet.Cells.Item(7,8).Text)"  | Tee-Object -Append -FilePath $ScreenOutput

    # Adjust row height based the information
    # Normal line height is 14.4
    $NumLines = ($OutputSheet.Cells.Item(6,10).Text).Split("`n").Count
    if ($NumLines -gt 1) {
        Write-Output "      Associated Documentation - Height adjusted" | Tee-Object -Append -FilePath $ScreenOutput
        $OutputSheet.Cells.Item(6,10).RowHeight = (14.4 * $NumLines)
    }

    $RiskAppliedToSomeone = $false
    Write-Output "    The risks apply to:" | Tee-Object -Append -FilePath $ScreenOutput
    try {
        if ($InputSheet.Shapes("Check Box 13").ControlFormat.Value -eq 1) {
            $RiskAppliedToSomeone = $true
            Write-Output "      Employee" | Tee-Object -Append -FilePath $ScreenOutput
        }
    }
    catch {
        Write-Output "      **ERROR** Employee checkbox missing" | Tee-Object -Append -FilePath $ScreenOutput
    }
    try {
        if ($InputSheet.Shapes("Check Box 14").ControlFormat.Value -eq 1) {
            $RiskAppliedToSomeone = $true
            Write-Output "      Contractor" | Tee-Object -Append -FilePath $ScreenOutput
        }
    }
    catch {
        Write-Output "      **ERROR** Contractor checkbox missing" | Tee-Object -Append -FilePath $ScreenOutput
    }
    try {
        if ($InputSheet.Shapes("Check Box 15").ControlFormat.Value -eq 1) {
            $RiskAppliedToSomeone = $true
            Write-Output "      Visitor" | Tee-Object -Append -FilePath $ScreenOutput
        }
    }
    catch {
        Write-Output "      **ERROR** Visitor checkbox missing" | Tee-Object -Append -FilePath $ScreenOutput
    }
    if (!$RiskAppliedToSomeone) {
        Write-Output "      **ERROR** This risk isn't applied to an Employee, Contractor or Visitor" | Tee-Object -Append -FilePath $ScreenOutput
    }

    # Data from the Risk Assessment Matrix
    # Capturing 6 columns worth of data per row
    Write-Output "    Risk Assessment Matrix being copied" | Tee-Object -Append -FilePath $ScreenOutput
    Write-Output "    `"Existing Hazard/Risk`" entries being copied" | Tee-Object -Append -FilePath $ScreenOutput

    $SourceRow = 22    # Starting row within the source spreadsheet
    $TargetRow = 21
    $AddedRows = 0

	if ($InputSheet.Cells.Item($SourceRow,1).Text.Length -gt 0) {
		while ($InputSheet.Cells.Item($SourceRow,1).Text.Length -gt 0) {
			Write-Output "      Hazard/Risk:     $($InputSheet.Cells.Item($SourceRow,1).Text)"  | Tee-Object -Append -FilePath $ScreenOutput
			Write-Output "        Assess NC - S: $($InputSheet.Cells.Item($SourceRow,13).Text)" | Tee-Object -Append -FilePath $ScreenOutput
			Write-Output "        Assess NC - L: $($InputSheet.Cells.Item($SourceRow,14).Text)" | Tee-Object -Append -FilePath $ScreenOutput

			# Is an extra line needed?
			if ($SourceRow -gt 46) {
				$objRange = $OutputExcel.Range("A"+[string](46+$AddedRows)).EntireRow
				[void]$objRange.Insert($xlShiftDown)

				# Need to know this later on
				$AddedRows++

				# Populate the row with original settings
				$RecordedRow22 = $OutputSheet.Range("A22").EntireRow
				$RecordedRow22.Copy() | Out-Null
				$InsertedRow = $OutputSheet.Range("A"+[string](35+$AddedRows))
				$InsertedRow.PasteSpecial(13) | Out-Null

				# Ignore the error caused by the formula
				# $OutputSheet.Cells.Item($TargetRow,39).Locked = $false

				# Add missing Risk number
				# $OutputSheet.Cells.Item($TargetRow,1)  = [string]"=A$(44+$AddedRows)+1"

				# Correct the Conditional Formatting which is hardcoded based on the row we copied earlier
				$Range = $OutputSheet.Range("AJ$(45+$AddedRows)") # :AJ$(35+$AddedRows)")
				$ConditionalFormatting = $Range.FormatConditions
				$ConditionalFormatting.Item(1).Modify(2,9,"=AJ$(45+$AddedRows)<>R$(45+$AddedRows)")
			} else {
				# Record "Persons at Risk"
				# Employee
				try {
					if ($InputSheet.Shapes("Check Box 13").ControlFormat.Value -ne 1) {
						$OutputSheet.Cells.Item($TargetRow,15) = ""
					}
				}
				catch {
				}
				# Contractor
				try {
					if ($InputSheet.Shapes("Check Box 14").ControlFormat.Value -ne 1) {
						$OutputSheet.Cells.Item($TargetRow,16) = ""
					}
				}
				catch {
				}
				# Visitor
				try {
					if ($InputSheet.Shapes("Check Box 15").ControlFormat.Value -ne 1) {
						$OutputSheet.Cells.Item($TargetRow,17) = ""
					}
				}
				catch {
				}
			}
			$OutputSheet.Cells.Item($TargetRow,3)   = $InputSheet.Cells.Item($SourceRow,1)    # Existing Hazard/Risk
			$OutputSheet.Cells.Item($TargetRow,3).HorizontalAlignment = $Left

			# Remove leading an trailing spaces
			# $OutputSheet.Cells.Item($TargetRow,3).Text = ($OutputSheet.Cells.Item($TargetRow,3).Text).Trim()

			# Adjust row height based the information
			# Normal line height is 15
            # Column width is about 31 characters
            [Array]$CellLines = ($OutputSheet.Cells.Item($TargetRow,3).Text).Split("`n")
			[Double]$NumLines = $CellLines.Count
            
            # Increase the number of lines based on the length 
            $CellLines | % {if ($_.Length -ge 31) {$NumLines += [int][math]::Floor($_.Length / 31)}}

			if ($NumLines -gt 1) {
				Write-Output "      Harzard/Risk - Height adjusted for $($NumLines) lines" | Tee-Object -Append -FilePath $ScreenOutput
				$OutputSheet.Cells.Item($TargetRow,3).RowHeight = (15 * $NumLines)
			}

			# Adjust the Hazard/Risk entry to be centered
			$OutputSheet.Cells.Item($TargetRow,3).VerticalAlignment = $Middle

			$OutputSheet.Cells.Item($TargetRow,18)  = $InputSheet.Cells.Item($SourceRow,13)   # Assessment (no controls) S
			$OutputSheet.Cells.Item($TargetRow,19)  = $InputSheet.Cells.Item($SourceRow,14)   # Assessment (no controls) L

			# Increase the rows so the next time round is working with the next line
			$SourceRow++
			$TargetRow++
		}
    } else {
		Write-Output "      No existing Hazard/Risks found at row $($SourceRow)" | Tee-Object -Append -FilePath $ScreenOutput
		Write-Output "        section skipped" | Tee-Object -Append -FilePath $ScreenOutput
	}
	# Blank out the rest of the Persons at Risk
    Write-Output "    Blank out Persons at Risk for blank hazard lines" | Tee-Object -Append -FilePath $ScreenOutput
    while ($AddedRows -eq 0) {
        # Has the Further Control Measures section been reached?
        # In all the template files their is a blank row before the next section
        #try {
            if ($OutputSheet.Cells.Item($SourceRow,1).Text -like "NOTE:*") {
                break
            }
        #}
        #catch {}

        $OutputSheet.Cells.Item($TargetRow,15) = ""
        $OutputSheet.Cells.Item($TargetRow,16) = ""
        $OutputSheet.Cells.Item($TargetRow,17) = ""

        # Increase the rows so the next time round is working with the next line
        $SourceRow++
        $TargetRow++

        # Check for a runaway search
        if ($SourceRow -ge 1000) {
            break
        }
    }

    $SourceRow = 22    # Starting row within the source spreadsheet
    $TargetRow = 21
	if ($InputSheet.Cells.Item($SourceRow,1).Text.Length -gt 0) {
        Write-Output "    `"Control Measures`" being copied" | Tee-Object -Append -FilePath $ScreenOutput

		while ($InputSheet.Cells.Item($SourceRow,17).Text.Length -ne 0) {
			Write-Output "      Ctrl Measure:    $($InputSheet.Cells.Item($SourceRow,17).Text)" | Tee-Object -Append -FilePath $ScreenOutput
			Write-Output "        Assess WC - S: $($InputSheet.Cells.Item($SourceRow,32).Text)" | Tee-Object -Append -FilePath $ScreenOutput
			Write-Output "        Assess WC - L: $($InputSheet.Cells.Item($SourceRow,33).Text)" | Tee-Object -Append -FilePath $ScreenOutput

			$OutputSheet.Cells.Item($TargetRow,22)  = $InputSheet.Cells.Item($SourceRow,17)   # Control Measures
			$OutputSheet.Cells.Item($TargetRow,36)  = $InputSheet.Cells.Item($SourceRow,32)   # Assessment (with controls) S
			$OutputSheet.Cells.Item($TargetRow,37)  = $InputSheet.Cells.Item($SourceRow,33)   # Assessment (with controls) L

			# Remove leading an trailing spaces
			# $OutputSheet.Cells.Item($TargetRow,37).Text = ($OutputSheet.Cells.Item($TargetRow,37).Text).Trim()

			# Adjust row height based the information
			# Normal line height is 15
            # Column width is about 64 characters
            [Array]$CellLines = ($OutputSheet.Cells.Item($TargetRow,22).Text).Split("`n")
			[Double]$NumLines = $CellLines.Count
            
            # Increase the number of lines based on the length 
            $CellLines | % {if ($_.Length -ge 64) {$NumLines += [int][math]::Floor($_.Length / 64)}}

			$NewRowHeight = 15 * $NumLines
			if ($NewRowHeight -gt $OutputSheet.Cells.Item($TargetRow,3).RowHeight) {
				Write-Output "      Control Measures - Height adjusted for $($NumLines) lines" | Tee-Object -Append -FilePath $ScreenOutput
				$OutputSheet.Cells.Item($TargetRow,22).RowHeight = $NewRowHeight
			}

			# Adjust the Hazard/Risk entry to be centered
			$OutputSheet.Cells.Item($TargetRow,22).VerticalAlignment = $Middle

			# Increase the rows so the next time round is working with the next line
			$SourceRow++
			$TargetRow++
		}
    } else {
		Write-Output "      No Control Measures found at row $($SourceRow)" | Tee-Object -Append -FilePath $ScreenOutput
		Write-Output "        section skipped" | Tee-Object -Append -FilePath $ScreenOutput
	}

    if ($AddedRows -gt 0) {
        Write-Output "    $($AddedRows) extra Hazard lines added" | Tee-Object -Append -FilePath $ScreenOutput
    }

    # Starting from the $SourceRow we now need to search for the next sections as the row changes
    $SignatureStartingRow = $SourceRow
    $Unsigned=$false
    $Undated=$false

    # Assessed By
    $SourceRow = $SignatureStartingRow
    while ($InputSheet.Cells.Item($SourceRow,1).Text -notlike "ASSESSED*") {
        $SourceRow++

        # Check for a runaway search
        if ($SourceRow -ge 1000) {
            break
        }
    }
    Write-Output "    Assessed by: $($InputSheet.Cells.Item($SourceRow,5).Text)" | Tee-Object -Append -FilePath $ScreenOutput
    Write-Output "      on:        $($InputSheet.Cells.Item($SourceRow,32).Text)" | Tee-Object -Append -FilePath $ScreenOutput
    #$OutputSheet.Cells.Item(50 + $AddedRows,7)   = $InputSheet.Cells.Item($SourceRow,5)    # Name
    #$OutputSheet.Cells.Item(50 + $AddedRows,22)  = $InputSheet.Cells.Item($SourceRow,17)   # Position
    #$OutputSheet.Cells.Item(50 + $AddedRows,36)  = $InputSheet.Cells.Item($SourceRow,32)   # Date
    if ($InputSheet.Cells.Item($SourceRow,5).Text.Trim().Length -lt 5) {$Unsigned=$true}
    if ($InputSheet.Cells.Item($SourceRow,32).Text.Trim().Length -eq 0) {$Undated=$true}

    # Approved By
    $SourceRow = $SignatureStartingRow
    while ($InputSheet.Cells.Item($SourceRow,1).Text -notlike "APPROVED*") {
        $SourceRow++

        # Check for a runaway search
        if ($SourceRow -ge 1000) {
            break
        }
    }
    Write-Output "    Approved by: $($InputSheet.Cells.Item($SourceRow,5).Text)" | Tee-Object -Append -FilePath $ScreenOutput
    Write-Output "      on:        $($InputSheet.Cells.Item($SourceRow,32).Text)" | Tee-Object -Append -FilePath $ScreenOutput
    #$OutputSheet.Cells.Item(52 + $AddedRows,7)   = $InputSheet.Cells.Item($SourceRow,5)    # Name
    #$OutputSheet.Cells.Item(52 + $AddedRows,22)  = $InputSheet.Cells.Item($SourceRow,17)   # Position
    #$OutputSheet.Cells.Item(52 + $AddedRows,36)  = $InputSheet.Cells.Item($SourceRow,32)   # Date
    if ($InputSheet.Cells.Item($SourceRow,5).Text.Trim().Length -lt 5) {$Unsigned=$true}
    if ($InputSheet.Cells.Item($SourceRow,32).Text.Trim().Length -eq 0) {$Undated=$true}

    # Reviewed By
    $SourceRow = $SignatureStartingRow
    while ($InputSheet.Cells.Item($SourceRow,1).Text -notlike "REVIEWED*") {
        $SourceRow++

        # Check for a runaway search
        if ($SourceRow -ge 1000) {
            break
        }
    }
    Write-Output "    Reviewed by: $($InputSheet.Cells.Item($SourceRow,5).Text)" | Tee-Object -Append -FilePath $ScreenOutput
    Write-Output "      on:        $($InputSheet.Cells.Item($SourceRow,32).Text)" | Tee-Object -Append -FilePath $ScreenOutput
    #$OutputSheet.Cells.Item(52 + $AddedRows,7)   = $InputSheet.Cells.Item($SourceRow,5)    # Name
    #$OutputSheet.Cells.Item(52 + $AddedRows,22)  = $InputSheet.Cells.Item($SourceRow,17)   # Position
    #$OutputSheet.Cells.Item(52 + $AddedRows,36)  = $InputSheet.Cells.Item($SourceRow,32)   # Date
    if ($InputSheet.Cells.Item($SourceRow,5).Text.Trim().Length -lt 5) {$Unsigned=$true}
    if ($InputSheet.Cells.Item($SourceRow,32).Text.Trim().Length -eq 0) {$Undated=$true}

    # If any of the signatures are black then highlight it in the notes file
    if (!$RiskAppliedToSomeone -or $Unsigned -or $Undated) {
        "`nFile: $($File.Name)" >> $ProcessingNotesFile
        if (!$RiskAppliedToSomeone) {
            "    This risk isn't applied to an Employee, Contractor or Visitor" >> $ProcessingNotesFile

            Write-Output "    Noted as not being applied to an Employee, Contractor or Visitor" | Tee-Object -Append -FilePath $ScreenOutput
        }
        if ($Unsigned) {
            "    This original Hazard/Risk file had 1 or more missing signatures" >> $ProcessingNotesFile

            Write-Output "    Noted as having a missing signature(s)" | Tee-Object -Append -FilePath $ScreenOutput
        }
        if ($Undated) {
            "    This original Hazard/Risk file had 1 or more missing sign-off dates" >> $ProcessingNotesFile

            Write-Output "    Noted as having a missing sign-off date(s)" | Tee-Object -Append -FilePath $ScreenOutput
        }
    }

	if ($PasswordProtectionRequired) {
		# Lock the template file again
		$InputSheet.Protect($SpreadsheetPassword)
	}

    # Save the file
    $OutputWorkBook.SaveAs($OutputFileName)
    Write-Output "  Saved:       $($OutputFileName)" | Tee-Object -Append -FilePath $ScreenOutput

    # Close the Spreadsheets
    $InputExcel.Quit()
    $OutputExcel.Quit()

	# Important: remove the used COM objects from memory
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$InputExcel)     | Out-Null
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$InputWorkBook)  | Out-Null
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$InputSheet)     | Out-Null

	[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$OutputExcel)    | Out-Null
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$OutputWorkBook) | Out-Null
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$OutputSheet)    | Out-Null

	# Belt & Braces to kill any Excel processes that might be running
	Get-Process -Name Excel | Stop-Process -Force
}

# Safty net to clear unused memory
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

$Finished = Get-Date
Write-Output "`n`n`n`n**************`n* STATISTICS *`n**************`n" | Tee-Object -Append -FilePath $ScreenOutput
Write-Output "Started:       $($Started.DateTime)" | Tee-Object -Append -FilePath $ScreenOutput
Write-Output "Finished:      $($Finished.DateTime)`n" | Tee-Object -Append -FilePath $ScreenOutput

$Duration = $Finished - $Started
Write-Output "Processed $($InputFileCount) files taking:" | Tee-Object -Append -FilePath $ScreenOutput
Write-Output "     Days: $($Duration.Days)" | Tee-Object -Append -FilePath $ScreenOutput
Write-Output "    Hours: $($Duration.Hours)" | Tee-Object -Append -FilePath $ScreenOutput
Write-Output "  Minutes: $($Duration.Minutes)" | Tee-Object -Append -FilePath $ScreenOutput
Write-Output "  Seconds: $($Duration.Seconds)" | Tee-Object -Append -FilePath $ScreenOutput
