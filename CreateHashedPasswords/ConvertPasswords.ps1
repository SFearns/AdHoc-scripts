##################################################################
## Copyright 2024  Stephen Fearns
##################################################################


# Load the PowerShell SQLite module
Import-Module DSInternals
Import-Module PSSQLite

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

    .PARAMETER OutputFile
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

    .PARAMETER ProgressBar
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$TRUE

    .PARAMETER OutputDB
		This switch is either $TRUE or $FALSE.

		$TRUE  will output the data into a CSV file
		$FALSE will output the data into a SQLite Database

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
		[string]$OutputFile = "HashedPasswords.SQLite",
		[switch]$OutputDB = $TRUE,
		[switch]$Overwrite = $FALSE,
		[switch]$LMHash = $TRUE,
		[switch]$NTHash = $TRUE,
		[switch]$ProgressBar = $TRUE
	)

	# When did the task start?
	$Started = Get-Date
	"Started: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Started.Year, $Started.Month, $Started.Day, $Started.Hour, $Started.Minute, $Started.Second

	# How large is the InputFile
	$InputFileSize = (Get-ChildItem $InputFile).Length
	"Filesize for '$($InputFile)': {0:n} MB`n" -f ($InputFileSize/1MB)

	# Remove '.\' from the beginning of the line
	if ($OutputFile.StartsWith('.\')) {$OutputFile = $OutputFile.Substring(2)}

	if ($OutputDB) {
		# Using a SQLite Database

		# Define The SQLiteDB filename variable
		if ($OutputFile.Substring(1).StartsWith(":\")) {
			# Starts with a drive letter
			$SQLiteDB = "$($OutputFile)"
		} else {
			$SQLiteDB = ".\$($OutputFile)"
		}

		# Does the Database need to be created?
		if (($TRUE -eq $Overwrite) -or !(Test-Path $OutputFile)) {
            # Remove the Database if it already exists and hide the output
            if (Test-Path $OutputFile) {
                "Removed: .\$($OutputFile)"
                Remove-Item -Path $OutputFile -Force | Out-Null
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
			"Using existing SQLite Database: $($OutputFile)"
		}

		# Make sure the Database exists
		if (!(Test-Path $OutputFile)) {
			"ERROR: Missing Database $($OutputFile)"
		}
	} else {
		# Using a CSV file

		# Change the output filename as a Database is being used
		$OutputFile = ".\$($InputFile).csv"

		# Are we writing the output to a fresh CSV file?
		if ((!$Overwrite) -or (!Test-Path $OutputFile)) {
			Write-Host "Created: $($OutputFile)"

			# Create the CSV Header row
			"Password,LMHash,NTHash" > $OutputFile
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

		# Update the progress variables
		$PasswordsProgressed++

		# Show the progress bar if required
		if ($ProgressBar) {
			$BytesProcessed += $Password.Length

			$PercentageCompleted = ($BytesProcessed/$InputFileSize * 100)
			$StatusText = "Processed {0:n}% / {1:n} MB of {2:n} MB -- Password: {3}" -f $PercentageCompleted, ($BytesProcessed/1MB), ($InputFileSize/1MB), $($Password)
			Write-Progress -PercentComplete $PercentageCompleted -Activity "Processing passwords from $($InputFile)" -Status $StatusText
		}

		# Is the password already in the Database?
		if ($OutputDB) {
			if (Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM HashedPasswords WHERE Password='$($Password.Replace('"','\"'))'") {
				# Record found
				"  Exists: $($Password)" >> $PasswordLog
				$SkipEntry = $TRUE
			}
		}

		if ($FALSE -eq $SkipEntry) {
			# Encrypt the clear text password for the Hash functions
			try {
				$SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

				# If the password is outside the 0-14 character range it can throw an error so
				# we shall hide those errors
				try {
					if ($NTHash) {$NTHashCode = ConvertTo-NTHash -Password $SecurePassword}
				}
				catch {
					"  ERROR: NTHash couldn't be produced for '$($Password)'"
					$NTHashError = $TRUE
				}
				try {
					if ($LMHash) {$LMHashCode = ConvertTo-LMHash -Password $SecurePassword}
				}
				catch {
					"  ERROR: LMHash couldn't be produced for '$($Password)'"
					$LMHashError = $TRUE
				}

				# Change the password to stop special character processing by SQL
				$SafePassword = $Password.Replace('\','\\').Replace("'","\`'").Replace(';','\;').Replace('--','\--').Replace('/*','\/*').Replace('*/','\*/').Replace('if','\if').Replace('else','\else').Replace('IF','\IF').Replace('ELSE','\ELSE').Replace('0x','\0x').Replace('+','\+').Replace('FROM','\FROM').Replace('from','\from').Replace('select','\select').Replace('SELECT','\SELECT')

				# If both hash function produced an error then doesn't produce an output
				if (!$NTHashError -and !$LMHashError)	{
					if ($OutputDB) {
						# DB output
						$Query = "INSERT INTO HashedPasswords (Password, LMHash, NTHash) VALUES (`"{0}`", `"{1}`", `"{2}`")" -f $SafePassword.Replace('"','""'),$LMHashCode.Replace('"','""'),$NTHashCode.Replace('"','""')

						try {
							Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
						}
						catch {
							throw "ERROR: Unable to add record for $($Password)"
						}			
					} else {
						# CSV output
						"`"$($Password)`",`"$($LMHashCode)`",`"$($NTHashCode)`"" >> $OutputFile
					}
				}
			}
			catch {
				"  ERROR: Invalid password '$($Password)'"
			}
		}
	}

	"`nPasswords processed: {0}`n" -f $PasswordsProgressed

	# When did the task finish?
	$Finished = Get-Date

	# How long did the work take?
	$HowLong = $Finished - $Started
	"Finished: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Finished.Year, $Finished.Month, $Finished.Day, $Finished.Hour, $Finished.Minute, $Finished.Second
	"Duration: {0:d2}d {1:d2}h {2:d2}m {3:d2}s`n" -f $HowLong.Days, $HowLong.Hours, $HowLong.Minutes, $HowLong.Seconds
}