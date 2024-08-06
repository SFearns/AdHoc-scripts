##################################################################
## Copyright 2024  Stephen Fearns
##################################################################

# Load the PowerShell SQLite module
Import-Module PSSQLite
Import-Module DSInternals

function Convert-Passwords {
<#
    .SYNOPSIS
        This function converts passwords found in a file to LMHASH and NTHASH values.

    .DESCRIPTION
        This function opens the Input file and reads a single line at a time and converts the contents to an LMHASH and NTHASH value.

		These values are then saved into a CSV file.

		There are 5 parameters of which the -InputFile is a requirement.

		This script depends on:
		  DSInternals from https://github.com/MichaelGrafnetter/DSInternals
	      PSSQLite    from https://github.com/RamblingCookieMonster/PSSQLite

    .PARAMETER InputFile
        The password file to process.
        
		Note:
		Each line is considered to be the password 

    .PARAMETER SQLiteDatabase
        This is the filename of the output file.

		If left blank then:
		  CSV    - The output file is the InputFile with '.csv' appended at the end
		  SQLite - The Database will be called 'HashedPasswords.SQLite'

    .PARAMETER Overwrite
		This switch is either $TRUE or $FALSE.

		$TRUE  will append the contents to an existing file - unless it doesn't exist
		$FALSE will create a new file, overwritting the file if it already exists

		Default value:	$FALSE

    .PARAMETER NTHASH
		This switch is either $TRUE or $FALSE.

		$TRUE  will produce an HASH value
		$FALSE will NOT produce an HASH value

		Default value:	$TRUE

    .PARAMETER LMHASH
		This switch is either $TRUE or $FALSE.

		$TRUE  will produce an HASH value
		$FALSE will NOT produce an HASH value

		Default value:	$TRUE

		Notes:
		The LMHASH function doesn't support Unicode characters and will produce an error

    .PARAMETER ShowProgressBar
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

    .PARAMETER OutputCSV
		This switch is either $TRUE or $FALSE.

		$TRUE  will output the data into a CSV file
		$FALSE will output the data into a SQLite Database

		Default value:	$FALSE

    .PARAMETER Verbose
		This switch is either $TRUE or $FALSE.

		Will work in the normal way but not fully implemented yet

		Default value:	$FALSE

    .INPUTS
        Piped values are not supported.

    .OUTPUTS
		The function uses a progress bar by default
		Progress information is output to the screen (which can be re-directed)

    .EXAMPLE
        Convert_passwords_to_hashes.ps1 -InputFile "passwords.txt"

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party module isn't possible

#>

	Param (
		[string]$InputFile = $(throw "-InputFile is required."),
		[string]$SQLiteDatabase = "HashedPasswords.SQLite",
		[switch]$Verbose = $FALSE,
		[switch]$OutputCSV = $FALSE,
		[switch]$Overwrite = $FALSE,
		[switch]$LMHash = $TRUE,
		[switch]$NTHash = $TRUE,
		[switch]$LogFile = $FALSE,
		[switch]$ShowProgressBar = $FALSE
	)

	# Record old Verbose setting
	if ($Verbose) {
		$OldVerbose = $VerbosePreference
		$VerbosePreference = "Continue"
	}

	# When did the task start?
	$Started = Get-Date
	"Started: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Started.Year, $Started.Month, $Started.Day, $Started.Hour, $Started.Minute, $Started.Second

	# Password Log filename
	if ($TRUE -eq $LogFile) {
		$PasswordLog = "$($InputFile).log"
		if (Test-Path $PasswordLog) {
			if ($PasswordLog.StartsWith('.\')) {$PasswordLog = $PasswordLog.Substring(2)}

			"Removed: .\$($PasswordLog)"
			Remove-Item -Path $PasswordLog -Force | Out-Null
		}
	}

	# How large is the InputFile
	$InputFileSize = (Get-ChildItem $InputFile).Length
	"Filesize for '$($InputFile)': {0:n} MB`n" -f ($InputFileSize/1MB)

	# Remove '.\' from the beginning of the line
	if ($SQLiteDatabase.StartsWith('.\')) {$SQLiteDatabase = $SQLiteDatabase.Substring(2)}

	if (!$OutputCSV) {
		# Using a SQLite Database

		# Define The SQLiteDB filename variable
		if ($SQLiteDatabase.Substring(1).StartsWith(":\")) {
			# Starts with a drive letter & folder
			$SQLiteDB = "$($SQLiteDatabase)"
		} else {
			# Must start with .\
			$SQLiteDB = ".\$($SQLiteDatabase)"
		}

		# Does the Database need to be created?
		if (($TRUE -eq $Overwrite) -or !(Test-Path $SQLiteDatabase)) {
            # Remove the Database if it already exists and hide the output
            if (Test-Path $SQLiteDatabase) {
                "Removed: $($SQLiteDatabase)"
                Remove-Item -Path $SQLiteDatabase -Force | Out-Null
            }

			# Create the tables with the required fields
			$Query = 'CREATE TABLE "HashedPasswords" ("ID" INTEGER NOT NULL UNIQUE, "Password" TEXT, "LMHash" TEXT KEY, "NTHash" TEXT KEY, PRIMARY KEY("ID" AUTOINCREMENT));'
			try {
				Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
                "Created: $($SQLiteDB) - Table"
			}
			catch {throw "ERROR: Unable to create $($SQLiteDB)"}	

			# Create the index for the ID field
			$Query = 'CREATE UNIQUE INDEX "ID" ON "HashedPasswords" ("ID" ASC);'
			try {
				Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
                "Created: $($SQLiteDB) - Index for ID"
			}
			catch {throw "ERROR: Unable to UNIQUE Index for ID -- $($SQLiteDB)"}	

			# Create a UNIQUE Index for the clear-text password
			$Query = 'CREATE UNIQUE INDEX "Password" ON HashedPasswords ("Password" ASC)'
			try {
				Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
                "Created: $($SQLiteDB) - Index for Password"
			}
			catch {throw "ERROR: Unable to UNIQUE Index for Password -- $($SQLiteDB)"}	

			# Create a UNIQUE Index for the clear-text password
			$Query = 'CREATE INDEX "NTHash" ON HashedPasswords ("NTHash" ASC)'
			try {
				Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
                "Created: $($SQLiteDB) - Index for NTHash"
			}
			catch {throw "ERROR: Unable to Index for NTHash -- $($SQLiteDB)"}	
			
			# Create a UNIQUE Index for the clear-text password
			$Query = 'CREATE INDEX "LMHash" ON HashedPasswords ("LMHash" ASC)'
			try {
				Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
                "Created: $($SQLiteDB) - Index for LMHash"
			}
			catch {throw "ERROR: Unable to Index for LMHash -- $($SQLiteDB)"}	
		} else {
			"Using existing SQLite Database: $($SQLiteDatabase)"
		}

		# Make sure the Database exists
		if (!(Test-Path $SQLiteDatabase)) {
			"ERROR: Missing Database $($SQLiteDatabase)"
		}
	} else {
		# Using a CSV file

		# Change the output filename as a Database is being used
		$SQLiteDatabase = ".\$($InputFile).csv"

		# Are we writing the output to a fresh CSV file?
		if ((!$Overwrite) -or (!Test-Path $SQLiteDatabase)) {
			Write-Host "Created: $($SQLiteDatabase)"

			# Create the CSV Header row
			"Password,LMHash,NTHash" > $SQLiteDatabase
		} else {
			"Using existing file: $(OutputFile)"
		}
	}

	# Used by the Status Bar
	$PasswordsProgressed = 0
	$BytesProcessed = 0

	# Work through each line of the folder file removing non ISO-8859-1 characters
	$InputFileWithPath = (Get-ChildItem $InputFile).FullName
	ForEach ($Password in [System.IO.File]::ReadLines($InputFileWithPath))
	{
		# Remove non ISO-8859-1 characters
		# $Password = $Password  -replace '\P{IsBasicLatin}'	# [^\p{IsBasicLatin}\p{IsLatin-1Supplement}]')
		# $Password = $Password  -replace '[^\p{IsBasicLatin}\p{IsLatin-1Supplement}]'
		$Password = $Password -replace '[^^\x30-\x39\x41-\x5A\x61-\x7A]+'

		# Reset temporary variables
		$SkipEntry   = $FALSE
		$NTHashError = $FALSE
		$LMHashError = $FALSE
		$NTHashCode  = ""
		$LMHashCode  = ""

		# Show the progress bar if required
		if ($ShowProgressBar) {
			$BytesProcessed += $Password.Length

			$PercentageCompleted = ($BytesProcessed/$InputFileSize * 100)
			$StatusText = "Processed {0:n}% / {1:n} MB of {2:n} MB -- Password: {3}" -f $PercentageCompleted, ($BytesProcessed/1MB), ($InputFileSize/1MB), $($Password)
			Write-Progress -PercentComplete $PercentageCompleted -Activity "Processing passwords from $($InputFile)" -Status $StatusText
		}

		# Update the progress variables
		$PasswordsProgressed++

		# Change the password to stop special characters being processing by SQL
		$SafePassword = $Password.Replace('\','\\').Replace("'","''").Replace(';','\;').Replace('--','\--').Replace('/*','\/*').Replace('*/','\*/').Replace('0x','\0x').Replace('+','\+')
		
		## .Replace('if','\if').Replace('else','\else').Replace('IF','\IF').Replace('ELSE','\ELSE').Replace('FROM','\FROM').Replace('from','\from').Replace('select','\select').Replace('SELECT','\SELECT')

		if ($TRUE -eq $LogFile) {"  Password changed from: $($Password)`n                     to: $($SafePassword)" >> $PasswordLog}

		# Is the password already in the Database?
		if (!$OutputCSV) {
			$SelectedRecord = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM HashedPasswords WHERE Password='$($SafePassword.Replace('"','\"'))'" -ErrorAction SilentlyContinue

			if ($SelectedRecord) {
				# Record found
				if ($TRUE -eq $LogFile) {"  Exists: $($Password) / $($SafePassword)" >> $PasswordLog}
				$SkipEntry = $TRUE
			}

			# Remove the temporary variable
			Remove-Variable -Name SelectedRecord
		}

		if ($FALSE -eq $SkipEntry) {
			# Encrypt the clear text password for the Hash functions
			try {
				$SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

				# If the password is outside the 0-14 character range it can throw an error so
				# we shall hide those errors
				try {
					if ($NTHash) {
						$NTHashCode = ConvertTo-NTHash -Password $SecurePassword -ErrorAction SilentlyContinue
					}
				}
				catch {
					Write-Verbose "  ERROR: NTHash couldn't be produced for '$($Password)'"
					$NTHashError = $TRUE
				}
				try {
					if ($LMHash) {
						$LMHashCode = ConvertTo-LMHash -Password $SecurePassword -ErrorAction SilentlyContinue
					}
				}
				catch {
					if ($Password.Length -gt 14) {
						Write-Verbose "  ERROR: LMHash couldn't be produced for '$($Password)' as it's >14 characters"
					} else {
						Write-Verbose "  ERROR: LMHash couldn't be produced for '$($Password)'"
					}
					$LMHashError = $TRUE
				}

				# If both hash function produced an error then doesn't produce an output
				if (!$NTHashError -or !$LMHashError)	{
					if (!$OutputCSV) {
						# DB output
						$Query = "INSERT INTO HashedPasswords (Password, LMHash, NTHash) VALUES ('{0}', '{1}', '{2}')" -f $SafePassword.Replace('"','\"'),$LMHashCode,$NTHashCode

						try {
							Invoke-SqliteQuery -DataSource $SQLiteDB -Query "$($Query)" -ErrorAction SilentlyContinue
						}
						catch {
							throw "ERROR: Unable to add record for '$($Password)' / '$($SafePassword.Replace('"','\"'))'"
						}			
					} else {
						# CSV output
						"`"$($Password)`",`"$($LMHashCode)`",`"$($NTHashCode)`"" >> $SQLiteDatabase
					}
				}
			}
			catch {
				Write-Verbose "  ERROR: Invalid password '$($Password)'"
			}
		}
	}

	# Reset the Verbose back to the original
	if ($Verbose) {
		$VerbosePreference = $OldVerbose
	}

	"`nPasswords processed: {0}`n" -f $PasswordsProgressed

	# When did the task finish?
	$Finished = Get-Date

	# How long did the work take?
	$HowLong = $Finished - $Started
	"Finished: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Finished.Year, $Finished.Month, $Finished.Day, $Finished.Hour, $Finished.Minute, $Finished.Second
	"Duration: {0:d2}d {1:d2}h {2:d2}m {3:d2}s`n" -f $HowLong.Days, $HowLong.Hours, $HowLong.Minutes, $HowLong.Seconds
}

function Find-Passwords {
<#
    .SYNOPSIS
        This function uses a SQLite Database to seach for LM & NT Hash values and if found displays the password(s).

    .DESCRIPTION
		Find the password that matches the Hash using a SQLite DB

		This script depends on:
		  DSInternals from https://github.com/MichaelGrafnetter/DSInternals
	      PSSQLite    from https://github.com/RamblingCookieMonster/PSSQLite

    .PARAMETER InputFile
        This file is the human readable dump file containing (in this order):
			Username
			UID
			LMHash
			NTHash
			<unknown>
			<unknown>

		Each field is seperated with a colon ( : )

    .PARAMETER SQLiteDatabase
        This is the SQLite Database to be used

    .PARAMETER ShowProgressBar
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

    .INPUTS
        Piped values are not supported.

    .OUTPUTS
		The function uses a progress bar by default
		Progress information is output to the screen (which can be re-directed)

    .EXAMPLE
        Find-Passwords -InputFile "sample-sam.txt" -SQLiteDatabase "HashedPasswords.SQLite" -ShowProgressBar

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party module isn't possible

#>

	Param (
		[string]$InputFile = $(throw "-InputFile is required."),
		[string]$SQLiteDatabase = $(throw "-SQLiteDatabase is required."),
		[switch]$ShowProgressBar = $FALSE
	)

	# When did the task start?
	$Started = Get-Date
	"Started: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Started.Year, $Started.Month, $Started.Day, $Started.Hour, $Started.Minute, $Started.Second

	# Load the -InputFile into memory
	[array]$InputFileContents = Get-Content -Path $InputFile

	"Passwords to process: {0}`n" -f $InputFileContents.Count

	# Remove '.\' from the beginning of the line
	if ($SQLiteDatabase.StartsWith('.\')) {$SQLiteDatabase = $SQLiteDatabase.Substring(2)}

	# Define The SQLiteDB filename variable
	if ($SQLiteDatabase.Substring(1).StartsWith(":\")) {
		# Starts with a drive letter & folder
		$SQLiteDB = "$($SQLiteDatabase)"
	} else {
		# Must start with .\
		$SQLiteDB = ".\$($SQLiteDatabase)"
	}

	# Make sure the Database exists
	if (!(Test-Path $SQLiteDatabase)) {
		"ERROR: Missing Database $($SQLiteDatabase)"
		Break
	}

	# Used by the Status Bar
	$PasswordsProgressed = 0

	# Work through each line of the file
	ForEach ($Line in $InputFileContents)
	{
		# Show the progress bar if required
		if ($ShowProgressBar) {
			$PercentageCompleted = ($PasswordsProgressed/$InputFileContents.Count * 100)
			$StatusText = "Processed {0:n}%" -f $PercentageCompleted
			Write-Progress -PercentComplete $PercentageCompleted -Activity "Processing passwords from $($InputFile)" -Status $StatusText
		}

		# Update the progress variables
		$PasswordsProgressed++

		# Break up the line into the seperate parts
		$LineParts = $Line.Split(':')

		# Is the password already in the Database?
		$FoundLMHash = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM HashedPasswords WHERE LMHash='$($LineParts[2])'"
		$FoundNTHash = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM HashedPasswords WHERE NTHash='$($LineParts[3])'"

		if (!($FoundLMHash)) {
			$FoundLMHash = ""
		} else {
			$FoundLMHash = $FoundLMHash.Password
		}
		if (!($FoundNTHash)) {
			$FoundNTHash = ""
		} else {
			$FoundNTHash = $FoundNTHash.Password
		}

		"$($LineParts[0]),$($LineParts[1]),$($FoundLMHash),$($FoundNTHash),$($LineParts[4]),$($LineParts[5])"
	}

	"`nPasswords processed: {0}`n" -f $PasswordsProgressed

	# When did the task finish?
	$Finished = Get-Date

	# How long did the work take?
	$HowLong = $Finished - $Started
	"Finished: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Finished.Year, $Finished.Month, $Finished.Day, $Finished.Hour, $Finished.Minute, $Finished.Second
	"Duration: {0:d2}d {1:d2}h {2:d2}m {3:d2}s`n" -f $HowLong.Days, $HowLong.Hours, $HowLong.Minutes, $HowLong.Seconds
}

function Import-Passwords {
<#
    .SYNOPSIS
        This function imports new passwords from a text file (1 password per line) and rejects duplicates

    .DESCRIPTION
		This function imports new passwords from a text file (1 password per line) and rejects duplicates.
		
		The NTHash and LMHash are NOT created as part of this process.  See 'Create-MissingHashes' for that function.

		This script depends on:
		  DSInternals from https://github.com/MichaelGrafnetter/DSInternals
	      PSSQLite    from https://github.com/RamblingCookieMonster/PSSQLite

    .PARAMETER PasswordFile
        This file containing the list of passowrds.  Each line is considered a password

    .PARAMETER SQLiteDatabase
        This is the SQLite Database to be used

    .PARAMETER ShowProgressBar
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

    .INPUTS
        Piped values are not supported.

    .OUTPUTS
		The function uses a progress bar by default
		Progress information is output to the screen (which can be re-directed)

    .EXAMPLE
        Import-Passwords -InputFile "passwords.txt" -SQLiteDatabase "HashedPasswords.SQLite" -ShowProgressBar:$true

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party module isn't possible

#>

	Param (
		[string]$InputFile = $(throw "-InputFile is required."),
		[string]$SQLiteDatabase = $(throw "-SQLiteDatabase is required."),
		[switch]$ShowProgressBar = $FALSE
	)

	Convert-Passwords -InputFile $InputFile -SQLiteDatabase $SQLiteDatabase -LMHash:$false -NTHash:$false -ShowProgressBar:$ShowProgressBar -LogFile:$false
}

function Create-Hashes {
<#
    .SYNOPSIS
        This function will create missing hashes in the SQL database

    .DESCRIPTION
		This function will create any missing hashes in the SQL database.  If creating a hash fails then a dummy value 'failed_to_create' will be entered.
		
		This script depends on:
		  DSInternals from https://github.com/MichaelGrafnetter/DSInternals
	      PSSQLite    from https://github.com/RamblingCookieMonster/PSSQLite

    .PARAMETER SQLiteDatabase
        This is the SQLite Database to be used

    .PARAMETER ShowProgressBar
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

    .INPUTS
        Piped values are not supported.

    .OUTPUTS
		The function uses a progress bar by default
		Progress information is output to the screen (which can be re-directed)

    .EXAMPLE
        Create-Hashes -SQLiteDatabase "HashedPasswords.SQLite" -ShowProgressBar:$true

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party module isn't possible

#>

	Param (
		[string]$SQLiteDatabase = $(throw "-SQLiteDatabase is required."),
		[switch]$ShowProgressBar = $FALSE
	)

	# Work through the database for blank hashes
}

function Find-ExcelPassword {
	Param (
		[string]$ExcelFile = $(throw "-ExcelFile is required."),
		[string]$SQLiteDatabase = $(throw "-SQLiteDatabase is required."),
		[switch]$ShowProgressBar = $FALSE
	)

	# When did the task start?
	$Started = Get-Date
	"Started: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Started.Year, $Started.Month, $Started.Day, $Started.Hour, $Started.Minute, $Started.Second

	# Remove '.\' from the beginning of the line
	if ($SQLiteDatabase.StartsWith('.\')) {$SQLiteDatabase = $SQLiteDatabase.Substring(2)}

	# Define The SQLiteDB filename variable
	if ($SQLiteDatabase.Substring(1).StartsWith(":\")) {
		# Starts with a drive letter & folder
		$SQLiteDB = "$($SQLiteDatabase)"
	} else {
		# Must start with .\
		$SQLiteDB = ".\$($SQLiteDatabase)"
	}

	# Make sure the Database exists
	if (!(Test-Path $SQLiteDatabase)) {
		"ERROR: Missing Database $($SQLiteDatabase)"
		Break
	}

	# Used by the Status Bar
	$PasswordsProgressed = 0

    # Need to work through the SQLite DB that contains the words to attempt
	# How many passwords are we working through
	$PasswordCount = (Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT seq FROM sqlite_sequence").seq

	# Get the  first password
	$Password = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT ID,Password FROM HashedPasswords WHERE ID=1"

	$ExitLoopReason = 0
	$PasswordFound = 1
	$BreakOnError = 2
	$EndOfDatabase = 3
	$ExcelObject  = New-Object -ComObject Excel.Application
	New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\$($ExcelObject.Version)\Excel\Security" -Name AccessVBOM -Value 1 -Force | Out-Null

	do {
		$PasswordsProgressed++

		# Update the Progress Bar if active
		if ($ShowProgressBar) {
			$Now = Get-Date
			$Minutes = ($Now - $Started).TotalMinutes
			$PasswordsPM = [Math]::Round($PasswordsProgressed / $Minutes)
			$StatusLine = "Passwords/m: {0:n0} -- {1:n0} of {2:n0} -- Password: {3}" -f $PasswordsPM, $Password.ID, $PasswordCount, $Password.Password
			Write-Progress -Activity "Finding password" -Status $StatusLine
		}

		# Do we have a valid password?
#		if (($Password -replace '\s', '').Count -gt 0) {
		if ($Password.Password.Length -gt 0) {
			# Attempt to open the spreadsheet using the password
			try {
				# Try and open the
				$WorkBook = $ExcelObject.Workbooks.Open($ExcelFile, $false, $false, [Type]::Missing, $Password.Password)

				# If we are still here the Open command worked
				$ExitLoopReason = $PasswordFound

				# Which password worked?
				"`nPassword found:      '{0}' on attempt {1:n0}" -f $Password.Password, $Password.ID
			}
			catch {
				# Check for an invalid password
				if ($PSItem.Exception.Message.StartsWith("The password you`'ve supplied is not correct.")) {
					# Couldn't open the file so try the next password
					try {
						# Read the next record if possible
						$Password = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT ID,Password FROM HashedPasswords WHERE ID = $($Password.ID + 1)"
					}
					catch {
						"`nPassword NOT Found: {0:n0} attempts" -f $Password.ID
						$ExitLoopReason = $EndOfDatabase
					}
				} else {
					if ($PSItem) { # .Exception.Message.StartsWith("Sorry, we couldn")) {
						"`n**************************************************`n"
						$PSItem.Exception.Message
						"`n**************************************************"
						$ExitLoopReason = $BreakOnError
						break
					}
				}
			}
		}
	} while ($ExitLoopReason -eq 0)

    # Close the Spreadsheets
    $ExcelObject.Quit()

	# Important: remove the used COM objects from memory
	[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$ExcelObject) | Out-Null
	if ($EXitLoopReason -ne $BreakOnError) {[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$WorkBook) | Out-Null}

	# Belt & Braces to kill any Excel processes that might be running
	Get-Process -Name Excel | Stop-Process -Force

	# Safty net to clear unused memory
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()

	if ($EXitLoopReason -ne $BreakOnError) {"`nPasswords processed: {0:n0}" -f $PasswordsProgressed}

	# When did the task finish?
	$Finished = Get-Date

	# How long did the work take?
	$HowLong = $Finished - $Started
	"`nFinished: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Finished.Year, $Finished.Month, $Finished.Day, $Finished.Hour, $Finished.Minute, $Finished.Second
	"Duration: {0:d2}d {1:d2}h {2:d2}m {3:d2}s`n" -f $HowLong.Days, $HowLong.Hours, $HowLong.Minutes, $HowLong.Seconds
}