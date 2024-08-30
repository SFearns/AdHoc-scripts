$Today=Get-Date
$AllStartDT=Get-Date

$CurrentDirectory='/'
# The following variable contains the path to the working folder
$HomeFolder='C:\WorkFolder'
# Directory listings are contained in the following sub-folder
$InputFolder=$HomeFolder+'\InputFiles'
# All the results go into the following sub-folder
$OutputFolder=$HomeFolder+'\Results'
# Summary file is named below
$ReportFile=$OutputFolder+'\Processing Report - Complete File breakdown.txt'
# The following file (within the working folder) contains the file extensions.
# Each entry will create its own CSV file
$FileExtensionFile=$HomeFolder+'\FileExtensions.csv'

$AllLineCounter=0

# Ranges were set by Arkivum
$Range1Min=0;      $Range1Max=1MB-1;    $Range1FName="Range1 (0MB to 1MB)"
$Range2Min=1MB;    $Range2Max=10MB-1;   $Range2FName="Range2 (1MB to 10MB)"
$Range3Min=10MB;   $Range3Max=50MB-1;   $Range3FName="Range3 (10MB to 50MB)"
$Range4Min=50MB;   $Range4Max=250MB-1;  $Range4FName="Range4 (50MB to 250MB)"
$RangeEEMin=250MB; $RangeEEMax=1TB;     $RangeEEFName="Range5 (EveryThing Else)"

# Set variables to ZERO
$Range1FileCounters=$Range2FileCounters=$Range3FileCounters=$Range4FileCounters=$RangeEEFileCounters=0
$Range1SizeCounters=$Range2SizeCounters=$Range3SizeCounters=$Range4SizeCounters=$RangeEESizeCounters=0
$Range1AllFileCounters=$Range2AllFileCounters=$Range3AllFileCounters=$Range4AllFileCounters=$RangeEEAllFileCounters=0
$Range1AllSizeCounters=$Range2AllSizeCounters=$Range3AllSizeCounters=$Range4AllSizeCounters=$RangeEEAllSizeCounters=0

# This function will compress a folder
function Compress-Archive([String]$Path,[String]$DestinationPath)
{
   Add-Type -Assembly System.IO.Compression.FileSystem
   $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
   [System.IO.Compression.ZipFile]::CreateFromDirectory($Path,$DestinationPath,$compressionLevel,$false)
}

Function Convert-Delimiter([regex]$from,[string]$to) 
{ 
   process
   {  
      ## replace the original delimiter with the new one, wrapping EVERY block in Þ
      ## if there's quotes around some text with a delimiter, assume it doesn't count
      ## if there are two quotes "" stuck together inside quotes, assume they're an 'escaped' quote
      $_ = $_ -replace "(?:`"((?:(?:[^`"]|`"`"))+)(?:`"$from|`"`$))|(?:((?:.(?!$from))*.)(?:$from|`$))","`$1`$2$to" 
      ## clean up the end where there might be duplicates
      $_ = $_ -replace "(?:$to|)?`$",""
      ## normalize quotes so that they're all double "" quotes
      $_ = $_ -replace "`"`"","`"" -replace "`"","`"`"" 
      ## remove the Þ wrappers if there are no quotes inside them
      $_ = $_ -replace "((?:[^`"](?!$to))+)($to|`$)","`$1`$2"
      ## replace the Þ with quotes, and explicitly emit the result
      write-output $_ # -replace "","`""
   }
}

Write-Output "Reading Directory ($InputFolder)"
$InputFolderDirectoryListing = Get-ChildItem -Path $InputFolder | Sort-Object Name

if ($InputFolderDirectoryListing) {
    Write-Host "Loading a list of file extensions"
    $FileExtensionTypes=Import-Csv -Path $FileExtensionFile

    $InputFolderDirectoryListing | ForEach-Object {
		# clear the screen - uncomment if required
		#Clear-Host

		# Set variables to ZERO for the processing of this file
		$LineCounter=0
        $Range1FileCounters=$Range2FileCounters=$Range3FileCounters=$Range4FileCounters=$RangeEEFileCounters=0
        $Range1SizeCounters=$Range2SizeCounters=$Range3SizeCounters=$Range4SizeCounters=$RangeEESizeCounters=0

        $DirListing=$InputFolder+'\'+$_.Name
        $StartDT=Get-Date; Write-Output "Processing:`t$DirListing`nStart DT:`t$StartDT"

		$PathLogFile=$OutputFolder+'\'+$_.Name
		# Comment out the next line if not required
		$PathNameDivision=$PathLogFile+"\Divisions"
		# Comment out the next line if not required
		$PathNameGroup=$PathLogFile+"\Groups"
		# Comment out the next line if not required
		$PathNameYears=$PathLogFile+"\Years"
		# Comment out the next line if not required
		$PathNameUsers=$PathLogFile+"\Users"
		# Comment out the next line if not required
		$PathNameType=$PathLogFile+"\FileType"
		# The following IS required
		$PathNameSize=$PathLogFile+"\FileSize"

		if ($PathLogFile      -and !(Test-Path -Path $PathLogFile))      {New-Item -Path $PathLogFile      -ItemType Directory | Out-Null}
		# Check the output folder could be created otherwise skip this input file.
		if (!(Test-Path -Path $PathLogFile))                             {Write-Output "ERROR: Can't create results sub-folder ($PathLogFile)"; break}

		if ($PathNameYears    -and !(Test-Path -Path $PathNameYears))    {New-Item -Path $PathNameYears    -ItemType Directory | Out-Null}
		if ($PathNameUsers    -and !(Test-Path -Path $PathNameUsers))    {New-Item -Path $PathNameUsers    -ItemType Directory | Out-Null}
		if ($PathNameDivision -and !(Test-Path -Path $PathNameDivision)) {New-Item -Path $PathNameDivision -ItemType Directory | Out-Null}
		if ($PathNameGroup    -and !(Test-Path -Path $PathNameGroup))    {New-Item -Path $PathNameGroup    -ItemType Directory | Out-Null}
		if ($PathNameSize     -and !(Test-Path -Path $PathNameSize))     {New-Item -Path $PathNameSize     -ItemType Directory | Out-Null}
		if ($PathNameType     -and !(Test-Path -Path $PathNameType))     {New-Item -Path $PathNameType     -ItemType Directory | Out-Null}

        ForEach ($Line in [System.IO.File]::ReadLines($DirListing)) {
            $LineCounter+=1
			# Is this line a directory line?
            if (($Line -like "/*:") -or ($Line -like "./*:") -or ($Line -like "*:")) {$Line=$CurrentDirectory=$Line.Replace('./','/').Replace(':','')}
            
			# Update the screen with the current line
			Write-Progress -Status "Line: $LineCounter" -Activity $CurrentDirectory -CurrentOperation $Line
			
            if (($PathNameDivision -or $PathNameGroup) -and $Line -like "/*") {
                # Process a Directory line
                $DivisionName=$GroupName=$null
                $DirectoryPath=$CurrentDirectory.Split('/')
				if ($DirectoryPath[5]) {$DivisionName=$DirectoryPath[5].Replace(' ','_').Replace('[','_').Replace(']','_').Replace('+','_')} else {$DivisionName=$null}
				if ($DirectoryPath[6]) {$GroupName=$DirectoryPath[6].Replace(' ','_').Replace('[','_').Replace(']','_').Replace('+','_')} else {$GroupName=$null}

                if ($PathNameDivision -and $DivisionName) {
                    $ttTempFile=$PathNameDivision+"\$($DivisionName).csv"
			        if (Test-Path -Path $ttTempFile) {$Line | Add-Content $ttTempFile} else {"path" | Set-Content $ttTempFile; $Line | Add-Content $ttTempFile}
                } else {
			        $DivisionName = 'Not found'
		        }
                if ($PathNameGroup -and $GroupName) {
                    $ttTempFile=$PathNameGroup+"\$($DivisionName) -- $($GroupName).csv"
			        if (Test-Path -Path $ttTempFile) {$Line | Add-Content $ttTempFile} else {"path" | Set-Content $ttTempFile; $Line | Add-Content $ttTempFile}
                }
	        }

            if ($Line -like "-*") {
                # Process a File line using RegEx
                $Line.Replace('ADDOMAIN\domain us','ADDOMAIN\domainuse').Replace('ADDOMAIN\domain ad','ADDOMAIN\domainadm') -match '^(?<Attributes>[rwx-]*)[ ]*[\+.][ ]*[\d]*[ ]*(?<UserID>[\d\w\\]*)[ ]*(?<Group>[\d\w\\]*)[ ]*(?<Size>\d*)[ ]*(?<Month>\w*)[ ]*(?<Day>\d*)[ ]*(?<Year_or_Time>[\d\:]*)[ ]*(?<FName>[\S\s]*)' | Out-Null
                $Time=$Day=$Month=$Year=$Size=$FName=$UserID=$null
                $UserID=$Matches.UserID
                [long]$Size=$Matches.Size
                $Month=$Matches.Month
                $Day=$Matches.Day
                if ($Matches.Year_or_Time -like "*:*") {
                    $Time=$Matches.Year_or_Time
                    if (([datetime]"$Month $Day" -ge [datetime]"Jan 01")-and([datetime]"$Month $Day"-le[datetime]"$($Today.Month) $($Today.Day)")) {$Year=$Today.Year} else {$Year=$Today.Year-1}
                } else {
                    $Year=$Matches.Year_or_Time
                }
				if ($PathNameType) {
					$FName=$Matches.FName
					$ttTempFile=$PathNameType+"\__No matched extension__.csv"
					if ($FName -Like "*.*") {
						$TempVariable=$FName.split('.')
						if ($FileExtensionTypes.Extension -contains $TempVariable[$TempVariable.count-1].ToLower()) {
							$ttTempFile=$PathNameType+"\$($TempVariable[$TempVariable.count-1].ToLower()).csv"
						}
					}
					if (Test-Path -Path $ttTempFile) {
						"$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
					} else {
						"owner,path,filename,year,month,day,time,bytesize" | Set-Content $ttTempFile
						"$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
					}
				}
				if ($PathNameYears) {
					$ttTempFile=$PathNameYears+"\$Year.csv"
					if (Test-Path -Path $ttTempFile) {
						"$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
					} else {
						"owner,path,filename,year,month,day,time,bytesize" | Set-Content $ttTempFile
						"$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
					}
				}
				if ($PathNameUsers) {
					$ttTempFile=$PathNameUsers+"\$($UserID.Replace('ADDOMAIN\','')).csv"
					if (Test-Path -Path $ttTempFile) {
						"$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
					} else {
						"owner,path,filename,year,month,day,time,bytesize" | Set-Content $ttTempFile
						"$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
					}
				}
                if (($Size -ge $Range1Min) -and ($Size -le $Range1Max)) {
                    $Range1FileCounters++
				    $Range1SizeCounters+=[long]$Size
				    $ttTempFile="$PathNameSize\$Range1FName.csv"
				    if (Test-Path -Path $ttTempFile) {
					    "$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
				    } else {
					    "owner,path,filename,year,month,day,time,bytesize" | Set-Content $ttTempFile
					    "$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
				    }
                } else {
                    if (($Size -ge $Range2Min) -and ($Size -le $Range2Max)) {
                        $Range2FileCounters++
				        $Range2SizeCounters+=[long]$Size
				        $ttTempFile="$PathNameSize\$Range2FName.csv"
				        if (Test-Path -Path $ttTempFile) {
					        "$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
				        } else {
					        "owner,path,filename,year,month,day,time,bytesize" | Set-Content $ttTempFile
					        "$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
				        }
                    } else {
                        if (($Size -ge $Range3Min) -and ($Size -le $Range3Max)) {
                            $Range3FileCounters++
				            $Range3SizeCounters+=[long]$Size
				            $ttTempFile="$PathNameSize\$Range3FName.csv"
				            if (Test-Path -Path $ttTempFile) {
					            "$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
				            } else {
					            "owner,path,filename,year,month,day,time,bytesize" | Set-Content $ttTempFile
					            "$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
				            }
                        } else {
                            if (($Size -ge $Range4Min) -and ($Size -le $Range4Max)) {
                                $Range4FileCounters++
				                $Range4SizeCounters+=[long]$Size
				                $ttTempFile="$PathNameSize\$Range4FName.csv"
				                if (Test-Path -Path $ttTempFile) {
					                "$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
				                } else {
					                "owner,path,filename,year,month,day,time,bytesize" | Set-Content $ttTempFile
					                "$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
				                }
                            } else {
                                $RangeEEFileCounters++
				                $RangeEESizeCounters+=[long]$Size
				                $ttTempFile="$PathNameSize\$RangeEEFName.csv"
				                if (Test-Path -Path $ttTempFile) {
					                "$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
				                } else {
					                "owner,path,filename,year,month,day,time,bytesize" | Set-Content $ttTempFile
					                "$($UserID.Replace('ADDOMAIN\','')),$CurrentDirectory,$FName,$Year,$Month,$Day,$Time,$Size" | Add-Content $ttTempFile
				                }
                            }
                        }
                    }
                }
            }
        }
		# Compress the Results
		Write-Output "Compressing the results"
        if (Test-Path -Path "$HomeFolder\Finished processing - $($_.BaseName).zip") {Remove-Item -Path "$HomeFolder\Finished processing - $($_.BaseName).zip"}
		Compress-Archive -Path $OutputFolder -DestinationPath "$HomeFolder\Finished processing - $($_.BaseName).zip"

        $AllLineCounter+=$LineCounter

        $Range1AllFileCounters+=$Range1FileCounters
        $Range2AllFileCounters+=$Range2FileCounters
        $Range3AllFileCounters+=$Range3FileCounters
        $Range4AllFileCounters+=$Range4FileCounters
        $RangeEEAllFileCounters+=$RangeEEFileCounters

        $Range1AllSizeCounters+=$Range1SizeCounters
        $Range2AllSizeCounters+=$Range2SizeCounters
        $Range3AllSizeCounters+=$Range3SizeCounters
        $Range4AllSizeCounters+=$Range4SizeCounters
        $RangeEEAllSizeCounters+=$RangeEESizeCounters

        $FinishDT=Get-Date
		"-------------------------------------------------------"
        "Processed file:`t$DirListing`n"
        "Start DT:`t$StartDT"
        "End DT:`t`t$FinishDT"
        "LineCount:`t{0:N0}" -f $LineCounter
        "Duration:`t$($FinishDT.Subtract($StartDT)) (d.hh:mm:ss.ms)"

        # Update the report file
        Write-Output "-------------------------------------------------------" | Add-Content $ReportFile
        "Processed file:`t`t$DirListing`n" | Add-Content $ReportFile
        "Start DT:`t`t`t$StartDT" | Add-Content $ReportFile
        "End DT:`t`t`t`t$FinishDT" | Add-Content $ReportFile
        "LineCount:`t`t`t{0:N0}" -f $LineCounter | Add-Content $ReportFile
        "Duration:`t`t`t$($FinishDT.Subtract($StartDT)) (d.hh:mm:ss.ms)`n" | Add-Content $ReportFile

        "$Range1FName :" | Add-Content $ReportFile
        "`tCount:`t`t`t{0:N0}" -f $Range1FileCounters | Add-Content $ReportFile
        "`tTotal Size:`t`t{0:N2} MB" -f ($Range1SizeCounters / 1MB) | Add-Content $ReportFile
        if ($Range1SizeCounters -gt 1024) {
            if (($Sum = (($Range1SizeCounters / $Range1FileCounters) / 1MB)) -lt 1) {
                "`tAvg Size:`t`t{0:N2} KB" -f (($Range1SizeCounters / $Range1FileCounters) / 1KB) | Add-Content $ReportFile
            } else {
                "`tAvg Size:`t`t{0:N2} MB" -f $Sum | Add-Content $ReportFile
            }
        } else {
            "`tAvg Size:`t`t0 MB" | Add-Content $ReportFile
        }
        "$Range2FName :" | Add-Content $ReportFile
        "`tCount:`t`t`t{0:N0}" -f $Range2FileCounters | Add-Content $ReportFile
        "`tTotal Size:`t`t{0:N2} MB" -f ($Range2SizeCounters / 1MB) | Add-Content $ReportFile
        if ($Range2SizeCounters -gt 1024) {
            if (($Sum = (($Range2SizeCounters / $Range2FileCounters) / 1MB)) -lt 1) {
                "`tAvg Size:`t`t{0:N2} KB" -f (($Range2SizeCounters / $Range2FileCounters) / 1KB) | Add-Content $ReportFile
            } else {
                "`tAvg Size:`t`t{0:N2} MB" -f $Sum | Add-Content $ReportFile
            }
        } else {
            "`tAvg Size:`t`t0 MB" | Add-Content $ReportFile
        }
        "$Range3FName :" | Add-Content $ReportFile
        "`tCount:`t`t`t{0:N0}" -f $Range3FileCounters | Add-Content $ReportFile
        "`tTotal Size:`t`t{0:N2} MB" -f ($Range3SizeCounters / 1MB) | Add-Content $ReportFile
        if ($Range3SizeCounters -gt 1024) {
            if (($Sum = (($Range3SizeCounters / $Range3FileCounters) / 1MB)) -lt 1) {
                "`tAvg Size:`t`t{0:N2} KB" -f (($Range3SizeCounters / $Range3FileCounters) / 1KB) | Add-Content $ReportFile
            } else {
                "`tAvg Size:`t`t{0:N2} MB" -f $Sum | Add-Content $ReportFile
            }
        } else {
            "`tAvg Size:`t`t0 MB" | Add-Content $ReportFile
        }
        "$Range4FName :" | Add-Content $ReportFile
        "`tCount:`t`t`t{0:N0}" -f $Range4FileCounters | Add-Content $ReportFile
        "`tTotal Size:`t`t{0:N2} MB" -f ($Range4SizeCounters / 1MB) | Add-Content $ReportFile
        if ($Range4SizeCounters -gt 1024) {
            if (($Sum = (($Range4SizeCounters / $Range4FileCounters) / 1MB)) -lt 1) {
                "`tAvg Size:`t`t{0:N2} KB" -f (($Range4SizeCounters / $Range4FileCounters) / 1KB) | Add-Content $ReportFile
            } else {
                "`tAvg Size:`t`t{0:N2} MB" -f $Sum | Add-Content $ReportFile
            }
        } else {
            "`tAvg Size:`t`t0 MB" | Add-Content $ReportFile
        }
        "$RangeEEFName :" | Add-Content $ReportFile
        "`tCount:`t`t`t{0:N0}" -f $RangeEEFileCounters | Add-Content $ReportFile
        "`tTotal Size:`t`t{0:N2} MB" -f ($RangeEESizeCounters / 1MB) | Add-Content $ReportFile
        if ($RangeEESizeCounters -gt 1024) {
            if (($Sum = (($RangeEESizeCounters / $RangeEEFileCounters) / 1MB)) -lt 1) {
                "`tAvg Size:`t`t{0:N2} KB" -f (($RangeEESizeCounters / $RangeEEFileCounters) / 1KB) | Add-Content $ReportFile
            } else {
                "`tAvg Size:`t`t{0:N2} MB" -f $Sum | Add-Content $ReportFile
            }
        } else {
            "`tAvg Size:`t`t0 MB" | Add-Content $ReportFile
        }
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

"$Range1FName :" | Add-Content $ReportFile
"`tCount:`t`t`t{0:N0}" -f $Range1AllFileCounters | Add-Content $ReportFile
"`tTotal Size:`t`t{0:N2} MB" -f ($Range1AllSizeCounters / 1MB) | Add-Content $ReportFile
if ($Range1AllSizeCounters -gt 1024) {
    if (($Sum = (($Range1AllSizeCounters / $Range1AllFileCounters) / 1MB)) -lt 1) {
        "`tAvg Size:`t`t{0:N2} KB" -f (($Range1AllSizeCounters / $Range1AllFileCounters) / 1KB) | Add-Content $ReportFile
    } else {
        "`tAvg Size:`t`t{0:N2} MB" -f $Sum | Add-Content $ReportFile
    }
} else {
    "`tAvg Size:`t`t0 MB" | Add-Content $ReportFile
}
"$Range2FName :" | Add-Content $ReportFile
"`tCount:`t`t`t{0:N0}" -f $Range2AllFileCounters | Add-Content $ReportFile
"`tTotal Size:`t`t{0:N2} MB" -f ($Range2AllSizeCounters / 1MB) | Add-Content $ReportFile
if ($Range2AllSizeCounters -gt 1024) {
    if (($Sum = (($Range2AllSizeCounters / $Range2AllFileCounters) / 1MB)) -lt 1) {
        "`tAvg Size:`t`t{0:N2} KB" -f (($Range2AllSizeCounters / $Range2AllFileCounters) / 1KB) | Add-Content $ReportFile
    } else {
        "`tAvg Size:`t`t{0:N2} MB" -f $Sum | Add-Content $ReportFile
    }
} else {
    "`tAvg Size:`t`t0 MB" | Add-Content $ReportFile
}
"$Range3FName :" | Add-Content $ReportFile
"`tCount:`t`t`t{0:N0}" -f $Range3AllFileCounters | Add-Content $ReportFile
"`tTotal Size:`t`t{0:N2} MB" -f ($Range3AllSizeCounters / 1MB) | Add-Content $ReportFile
if ($Range3AllSizeCounters -gt 1024) {
    if (($Sum = (($Range3AllSizeCounters / $Range3AllFileCounters) / 1MB)) -lt 1) {
        "`tAvg Size:`t`t{0:N2} KB" -f (($Range3AllSizeCounters / $Range3AllFileCounters) / 1KB) | Add-Content $ReportFile
    } else {
        "`tAvg Size:`t`t{0:N2} MB" -f $Sum | Add-Content $ReportFile
    }
} else {
    "`tAvg Size:`t`t0 MB" | Add-Content $ReportFile
}
"$Range4FName :" | Add-Content $ReportFile
"`tCount:`t`t`t{0:N0}" -f $Range4AllFileCounters | Add-Content $ReportFile
"`tTotal Size:`t`t{0:N2} MB" -f ($Range4AllSizeCounters / 1MB) | Add-Content $ReportFile
if ($Range4AllSizeCounters -gt 1024) {
    if (($Sum = (($Range4AllSizeCounters / $Range4AllFileCounters) / 1MB)) -lt 1) {
        "`tAvg Size:`t`t{0:N2} KB" -f (($Range4AllSizeCounters / $Range4AllFileCounters) / 1KB) | Add-Content $ReportFile
    } else {
        "`tAvg Size:`t`t{0:N2} MB" -f $Sum | Add-Content $ReportFile
    }
} else {
    "`tAvg Size:`t`t0 MB" | Add-Content $ReportFile
}
"$RangeEEFName :" | Add-Content $ReportFile
"`tCount:`t`t`t{0:N0}" -f $RangeEEAllFileCounters | Add-Content $ReportFile
"`tTotal Size:`t`t{0:N2} MB" -f ($RangeEEAllSizeCounters / 1MB) | Add-Content $ReportFile
if ($RangeEEAllSizeCounters -gt 1024) {
    if (($Sum = (($RangeEEAllSizeCounters / $RangeEEAllFileCounters) / 1MB)) -lt 1) {
        "`tAvg Size:`t`t{0:N2} KB" -f (($RangeEEAllSizeCounters / $RangeEEAllFileCounters) / 1KB) | Add-Content $ReportFile
    } else {
        "`tAvg Size:`t`t{0:N2} MB" -f $Sum | Add-Content $ReportFile
    }
} else {
    "`tAvg Size:`t`t0 MB" | Add-Content $ReportFile
}
