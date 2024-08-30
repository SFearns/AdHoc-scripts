$Today=Get-Date
$AllStartDT=Get-Date

$CurrentDirectory='/'
# The following variable contains the path to the working folder
$HomeFolder='C:\WorkFolder'
# Directory listings are contained in the following sub-folder
$InputFolder=$HomeFolder+'\InputFiles'
# Summary file is named below
$ReportFile=$HomeFolder+'\Processed Report - Complete File breakdown.txt'

$AllLineCounter=0

Write-Output "Reading Directory ($InputFolder)"
$InputFolderDirectoryListing = Get-ChildItem -Path $InputFolder | Sort-Object Name

if ($InputFolderDirectoryListing) {
    $InputFolderDirectoryListing | ForEach-Object {
		# clear the screen - uncomment if required
		#Clear-Host

		# Set variables to ZERO for the processing of this file
		$LineCounter=0
        $CurrentDirectory='<SMB ROOT>'

        $DirFileName=$_.Name
        $DirListing=$InputFolder+'\'+$DirFileName
        $StartDT=Get-Date; Write-Output "Processing:`t$DirListing`nStart DT:`t$StartDT"

        ForEach ($Line in [System.IO.File]::ReadLines($DirListing)) {
            $LineCounter+=1
            
            # Process a File line using RegEx
            $Line.Replace('ADDOMAIN\domain us','ADDOMAIN\domainuse').Replace('ADDOMAIN\domain ad','ADDOMAIN\domainadm') -match '^(?<tenno>[\d]*)[ ]*(?<blockcount>[\d]*)[ ]*(?<Attributes>[rwx-]*)[ ]*[\d]*[ ]*(?<UserID>[-\d\w\\]*)[ ]*(?<Group>[-\d\w\\]*)[ ]*(?<Size>\d*)[ ]*(?<Month>\w*)[ ]*(?<Day>\d*)[ ]*(?<Year_or_Time>[\d\:]*)[ ]*(?<FullFileName>[\S\s]*)' | Out-Null
            $Time=$Day=$Month=$Year=$Size=$FName=$UserID=$null
            $UserID=$Matches.UserID
            [long]$Size=$Matches.Size
            $Month=$Matches.Month
            $Day=$Matches.Day
            $FullFileName=$Matches.FullFileName

            # Extract the path and filename into seperate variables
            $TempVariable=[array]$FullFileName.split('/')
            $FName=$TempVariable[$TempVariable.count-1].ToLower()
            $FPath=$FullFileName.Replace("/$FName",'')
            if ($FPath.length -eq 0) {$FPath='/'}

			# Update the screen with the current line
			Write-Progress -Status "Line: $LineCounter" -Activity $FPath -CurrentOperation $Line
			
            if ($Matches.Year_or_Time -like "*:*") {
                $Time=$Matches.Year_or_Time
                if (([datetime]"$Month $Day" -ge [datetime]"Jan 01")-and([datetime]"$Month $Day"-le[datetime]"$($Today.Month) $($Today.Day)")) {$Year=$Today.Year} else {$Year=$Today.Year-1}
            } else {
                $Year=$Matches.Year_or_Time
            }
            if ($FName -Like "*.*") {
				$TempVariable=$FName.split('.')
				$FType = $TempVariable[$TempVariable.count-1].ToLower()
			}
            # Output data into SQL
            $SQLCommand = "INSERT INTO [DirectoryListings].[dbo].[DirectoryEntry] (StorageName,SubmitDateTime,LogFile,Owner,Path,FileName,FileType,Size,CreationDateTime,SourceLine) VALUES ('{0}',CONVERT(Datetime,'{1}',103),'{2}','{3}','{4}','{5}','{6}','{7}',CONVERT(Datetime,'{8}',103),'{9}')" -f "Isilon",$Today,$DirFileName.Replace("'","''"),$UserID.Replace('ADDOMAIN\','').Replace("'","''"),$FPath.Replace("'","''"),$FName.Replace("'","''"),$FType.Replace("'","''"),$Size,[datetime]"$Year-$Month-$Day $Time",$Line.Replace("'","''")

            if (Invoke-Sqlcmd -Query $SQLCommand -ServerInstance "SQLServer" -Username "dbuser" -Password "It's4ComplexPassword" -QueryTimeout 10) {
                $SQLCommand
            }
        }

        $AllLineCounter+=$LineCounter

        $FinishDT=Get-Date
		"-------------------------------------------------------"
        "Processed file:`t$DirListing`n"
        "Start DT:`t$StartDT"
        "End DT:`t`t$FinishDT"
        "LineCount:`t{0:N0}" -f $LineCounter
        "Duration:`t$($FinishDT.Subtract($StartDT)) (d.hh:mm:ss.ms)"
		"-------------------------------------------------------"

        # Update the report file
        Write-Output "-------------------------------------------------------" | Add-Content $ReportFile
        "Processed file:`t`t$DirListing`n" | Add-Content $ReportFile
        "Start DT:`t`t`t$StartDT" | Add-Content $ReportFile
        "End DT:`t`t`t`t$FinishDT" | Add-Content $ReportFile
        "LineCount:`t`t`t{0:N0}" -f $LineCounter | Add-Content $ReportFile
        "Duration:`t`t`t$($FinishDT.Subtract($StartDT)) (d.hh:mm:ss.ms)`n" | Add-Content $ReportFile
        Write-Output "-------------------------------------------------------" | Add-Content $ReportFile

    }
}

"======================================================="
"FINAL STATS`n"
"Start DT:`t$AllStartDT"
"End DT:`t`t$FinishDT"
"LineCount:`t{0:N0}" -f $AllLineCounter
"Duration:`t$($FinishDT.Subtract($AllStartDT)) (d.hh:mm:ss.ms)"

# Update the report file
"=======================================================" | Add-Content $ReportFile
"FINAL STATS`n" | Add-Content $ReportFile
"Start DT:`t`t`t$AllStartDT" | Add-Content $ReportFile
"End DT:`t`t`t`t$FinishDT" | Add-Content $ReportFile
"LineCount:`t`t`t{0:N0}" -f $AllLineCounter | Add-Content $ReportFile
"Duration:`t`t`t$($FinishDT.Subtract($AllStartDT)) (d.hh:mm:ss.ms)`n" | Add-Content $ReportFile
