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

    .PARAMETER HideProgressBar
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
		[switch]$HideProgressBar = $FALSE
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

			$Query = "CREATE TABLE HashedPasswords (Password TEXT PRIMARY KEY, LMHash TEXT KEY, NTHash TEXT KEY)"

			try {
				Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
                "Created: $($SQLiteDB)"
			}
			catch {
				throw "ERROR: Unable to create $($SQLiteDB)"
			}			
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

	# Work through each line of the folder file
	$InputFileWithPath = (Get-ChildItem $InputFile).FullName
	ForEach ($Password in [System.IO.File]::ReadLines($InputFileWithPath))
	{
		# Reset temporary variables
		$SkipEntry   = $FALSE
		$NTHashError = $FALSE
		$LMHashError = $FALSE
		$NTHashCode  = ""
		$LMHashCode  = ""

		# Show the progress bar if required
		if (!$HideProgressBar) {
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

    .PARAMETER ProgressBar
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$TRUE

    .INPUTS
        Piped values are not supported.

    .OUTPUTS
		The function uses a progress bar by default
		Progress information is output to the screen (which can be re-directed)

    .EXAMPLE
        Find-Passwords -InputFile "sample-sam.txt" -SQLiteDatabase "HashedPasswords.SQLite"

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party module isn't possible

    #>

	Param (
		[string]$InputFile = $(throw "-InputFile is required."),
		[string]$SQLiteDatabase = $(throw "-SQLiteDatabase is required."),
		[switch]$HideProgressBar = $FALSE
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
		if (!$HideProgressBar) {
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