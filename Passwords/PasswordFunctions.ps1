##################################################################
## Copyright 2024  Stephen Fearns
##################################################################

# Load the PowerShell SQLite module
Import-Module PSSQLite
Import-Module DSInternals

$pfVersion = "v2024.10.06"

Write-Host "`nPassword Functions  $($pfVersion)"
Write-Host "`nList all available functions with: " -NoNewline
Write-Host "Get-pfCommands" -ForegroundColor Yellow

function Get-pfCommands {
    [CmdletBinding()]
    Param ()
	Write-Host "`nPassword Functions  $($pfVersion)"

	Write-Host "`nThe added commands are:"
	Write-Host "    Get-pfCommands            - Lists the commands added"
	Write-Host "    Set-SQLiteDatabase        - Create the SQLite Database"
	Write-Host "    Get-SQLSafeText           - Package the string with SQL escape characters where required"
	Write-Host "    Import-Passwords          - Import Passwords from a text file (1 password per line)"
	Write-Host "    Find-Passwords            - Find the password given a Hash"
	Write-Host "    Add-MissingData           - Add missing data to the database after new fields were added"
	Write-Host "    Add-MissingHashes         - Calls 'Add-MissingData' to add just the NT and LM Hashes"
	Write-Host "    Add-MissingPasswordLength - Calls 'Add-MissingData' to add the length of the password"
	Write-Host "    Find-ExcelPassword        - Attempts to open an Excel spreadsheet using all the passwords from the database"
	Write-Host "    Import-COMBPasswords      - Import Passwords from a text file (: seperated file)`n"
}


##########################
# Setup module variables #
##########################

# Values are made ReadOnly so they can be removed if required without having to reload the CLI
Set-Variable pfDigits      -Force -ErrorAction SilentlyContinue -Option ReadOnly -Value '[0-9]'
Set-Variable pfLowerCase   -Force -ErrorAction SilentlyContinue -Option ReadOnly -Value '[a-z]'
Set-Variable pfUpperCase   -Force -ErrorAction SilentlyContinue -Option ReadOnly -Value '[A-Z]'
Set-Variable pfSpecials    -Force -ErrorAction SilentlyContinue -Option ReadOnly -Value '[^a-zA-Z0-9]'
Set-Variable pfEmptyLMHash -Force -ErrorAction SilentlyContinue -Option ReadOnly -Value 'aad3b435b51404eeaad3b435b51404ee'

function Set-SQLiteDatabase {
	Param (
		[string]$SQLiteDatabase = "Passwords.SQLite"
	)

	# Create the tables with the required fields
	$Query = 'CREATE TABLE "Passwords" ("Password" TEXT, "ID" INTEGER NOT NULL UNIQUE, "PasswordLength" INTEGER, "LMHash" TEXT KEY, "NTHash" TEXT KEY, "LowerCase" INTEGER, "UpperCase" INTEGER, "Digits" INTEGER, "Specials" INTEGER, PRIMARY KEY("ID" AUTOINCREMENT)) STRICT;'

	try {
		Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
		"Created: $($SQLiteDB) - Table"
	}
	catch {throw "ERROR: Unable to create $($SQLiteDB)"}	

	# Create the index for the ID field
	$Query = 'CREATE UNIQUE INDEX "ID" ON "Passwords" ("ID" ASC);'
	try {
		Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
		"Created: $($SQLiteDB) - Index for ID"
	}
	catch {throw "ERROR: Unable to UNIQUE Index for ID -- $($SQLiteDB)"}	

	# Create a UNIQUE Index for the clear-text password
	$Query = 'CREATE UNIQUE INDEX "Password" ON Passwords ("Password" ASC)'
	try {
		Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
		"Created: $($SQLiteDB) - Index for Password"
	}
	catch {throw "ERROR: Unable to UNIQUE Index for Password -- $($SQLiteDB)"}	

	# Create a UNIQUE Index for the clear-text password
	$Query = 'CREATE INDEX "PasswordLength" ON Passwords ("PasswordLength" ASC)'
	try {
		Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
		"Created: $($SQLiteDB) - Index for PasswordLength"
	}
	catch {throw "ERROR: Unable to UNIQUE Index for Password -- $($SQLiteDB)"}	

	# Create a UNIQUE Index for the clear-text password
	$Query = 'CREATE INDEX "NTHash" ON Passwords ("NTHash" ASC)'
	try {
		Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
		"Created: $($SQLiteDB) - Index for NTHash"
	}
	catch {throw "ERROR: Unable to Index for NTHash -- $($SQLiteDB)"}	
	
	# Create a UNIQUE Index for the clear-text password
	$Query = 'CREATE INDEX "LMHash" ON Passwords ("LMHash" ASC)'
	try {
		Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
		"Created: $($SQLiteDB) - Index for LMHash"
	}
	catch {throw "ERROR: Unable to Index for LMHash -- $($SQLiteDB)"}	
	
	# Create a UNIQUE Index for the clear-text password
	$Query = 'CREATE INDEX "LowerCase" ON Passwords ("LowerCase" ASC)'
	try {
		Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
		"Created: $($SQLiteDB) - Index for LowerCase"
	}
	catch {throw "ERROR: Unable to Index for LowerCase -- $($SQLiteDB)"}	
	
	# Create a UNIQUE Index for the clear-text password
	$Query = 'CREATE INDEX "UpperCase" ON Passwords ("UpperCase" ASC)'
	try {
		Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
		"Created: $($SQLiteDB) - Index for UpperCase"
	}
	catch {throw "ERROR: Unable to Index for UpperCase -- $($SQLiteDB)"}	
	
	# Create a UNIQUE Index for the clear-text password
	$Query = 'CREATE INDEX "Digits" ON Passwords ("Digits" ASC)'
	try {
		Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
		"Created: $($SQLiteDB) - Index for Digits"
	}
	catch {throw "ERROR: Unable to Index for Digits -- $($SQLiteDB)"}	
	
	# Create a UNIQUE Index for the clear-text password
	$Query = 'CREATE INDEX "Specials" ON Passwords ("Specials" ASC)'
	try {
		Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query
		"Created: $($SQLiteDB) - Index for Specials"
	}
	catch {throw "ERROR: Unable to Index for Specials -- $($SQLiteDB)"}	
}

function Get-SQLSafeText {
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

    .PARAMETER Text
        The Text to make safe for SQL

    .INPUTS
        Piped values are not supported.

    .OUTPUTS
		The function uses a progress bar by default
		Progress information is output to the screen (which can be re-directed)

    .EXAMPLE
        Get-SafeSQLString -InputFile "text%goes'here"

		Will return the following text

			text%%goes''here

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party modules still needs work
#>

	Param (
		[string]$Text = $(throw "-Text is required.")
	)

	[string]$Output = ''
	[string]$Output = $Text.Replace('\','\\').Replace("'","''").Replace(';','\;').Replace('--','\--').Replace('/*','\/*').Replace('*/','\*/').Replace('0x','\0x').Replace('+','\+').Replace('%','%%').Replace('"','\"')
		
	## .Replace('if','\if').Replace('else','\else').Replace('IF','\IF').Replace('ELSE','\ELSE').Replace('FROM','\FROM').Replace('from','\from').Replace('select','\select').Replace('SELECT','\SELECT')

	return $Output
}

function Import-Passwords {
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
		  SQLite - The Database will be called 'Passwords.SQLite'

    .PARAMETER NTHASH
		This switch is either $TRUE or $FALSE.

		$TRUE  will produce an HASH value
		$FALSE will NOT produce an HASH value

		Default value:	$FALSE

    .PARAMETER LMHASH
		This switch is either $TRUE or $FALSE.

		$TRUE  will produce an HASH value
		$FALSE will NOT produce an HASH value

		Default value:	$FALSE

		Notes:
		The LMHASH function doesn't support Unicode characters and will produce an error

    .PARAMETER ShowProgressBar
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

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
        Convert-Passwords -InputFile "passwords.txt" -SQLiteDatabase "Passwords.SQLite" -ShowProgressBar

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party modules still needs work
#>

	Param (
		[string]$InputFile = $(throw "-InputFile is required."),
		[string]$SQLiteDatabase = "Passwords.SQLite",
		[switch]$Verbose = $false,
		[switch]$LMHash = $false,
		[switch]$NTHash = $false,
		[switch]$ShowProgressBar = $false
	)

	# Record old Verbose setting
	if ($Verbose) {
		$OldVerbose = $VerbosePreference
		$VerbosePreference = "Continue"
	}

	# When did the task start?
	$Started = Get-Date
	"Started: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Started.Year, $Started.Month, $Started.Day, $Started.Hour, $Started.Minute, $Started.Second

	# How large is the InputFile
	$InputFileSize = (Get-ChildItem $InputFile).Length
	if ($InputFileSize -lt 1GB) {$InputFileSizeStr = "{0:n} MB" -f ($InputFileSize/1MB)} else {$InputFileSizeStr = "{0:n} GB" -f ($InputFileSize/1GB)}
	"Filesize for '$($InputFile)': {0:n}`n" -f $InputFileSizeStr

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

	# Does the Database need to be created?
	if (!(Test-Path $SQLiteDatabase)) {
		# Remove the Database if it already exists and hide the output
		if (Test-Path $SQLiteDatabase) {
			"Removed: $($SQLiteDatabase)"
			Remove-Item -Path $SQLiteDatabase -Force | Out-Null
		}

		Set-SQLiteDatabase -SQLiteDatabase $SQLiteDB
	} else {
		"Using existing SQLite Database: $($SQLiteDatabase)"
	}

	# Make sure the Database exists
	if (!(Test-Path $SQLiteDatabase)) {
		"ERROR: Missing Database $($SQLiteDatabase)"
		return
	}

	# Used by the Status Bar
	$PasswordsProgressed = 0
	$PasswordsAdded = 0
	$BytesProcessed = 0
	if ($InputFileSize -lt 1GB) {$InputFileSizeStr = "{0:n} MB" -f ($InputFileSize/1MB)} else {$InputFileSizeStr = "{0:n} GB" -f ($InputFileSize/1GB)}

	# Work through each line of the folder file removing non ISO-8859-1 characters
	$InputFileWithPath = (Get-ChildItem $InputFile).FullName
	ForEach ($Password in [System.IO.File]::ReadLines($InputFileWithPath))
	{
		# Update the progress variables
		$PasswordsProgressed++

		# Remove non Latin characters
		# $Password = $Password -replace '\P{IsBasicLatin}'	# [^\p{IsBasicLatin}\p{IsLatin-1Supplement}]')
		$Password = $Password -replace '[^\p{IsBasicLatin}\p{IsLatin-1Supplement}]'
		# $Password = $Password -replace '[^^\x30-\x39\x41-\x5A\x61-\x7A]+'

		# Reset temporary variables
		$SkipEntry         = $FALSE
		$NTHashError       = $FALSE
		$LMHashError       = $FALSE
		$NTHashCode        = $pfEmptyLMHash
		$LMHashCode        = ""

		# Show the progress bar if required
		if ($ShowProgressBar) {
			$BytesProcessed += $Password.Length

			$PercentageCompleted = ($BytesProcessed/$InputFileSize * 100)
			if ($BytesProcessed -lt 1GB) {$BytesProcessedStr = "{0:n} MB" -f ($BytesProcessed/1MB)} else {$BytesProcessedStr = "{0:n} GB" -f ($BytesProcessed/1GB)}
			$StatusText = "Processed {0:n}% ({1} of {2}) -- Added {3:n0} of {4:n0} -- {5}" -f $PercentageCompleted, $BytesProcessedStr, $InputFileSizeStr, $PasswordsAdded, $PasswordsProgressed, $($Password)
			Write-Progress -PercentComplete $PercentageCompleted -Activity "Processing passwords from $($InputFile)" -Status $StatusText
		}

		# Change the password to stop special characters being processing by SQL
		$SafePassword = Get-SQLSafeText -Text $Password

		Write-Verbose "  Password changed from: $($Password)`n                     to: $($SafePassword)"

		# Is the password already in the Database?
		$SelectedRecord = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM Passwords WHERE Password='$($SafePassword)'" -ErrorAction SilentlyContinue

		if ($SelectedRecord) {
			# Record found
			Write-Verbose "  Exists: $($Password) / $($SafePassword)"

			# Remove the temporary variable
			Remove-Variable -Name SelectedRecord
		} else {
			# Encrypt the clear text password for the Hash functions
			try {
				$SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

				# If the password is outside the 0-14 character range it can throw an error so
				# we shall hide those errors
				if ($NTHash) {
					try {
						$NTHashCode = ConvertTo-NTHash -Password $SecurePassword -ErrorAction SilentlyContinue
					}
					catch {
						Write-Verbose "  ERROR: NTHash couldn't be produced for '$($Password)'"
						$NTHashError = $TRUE
					}
				}
				if ($LMHash) {
					try {
						$LMHashCode = ConvertTo-LMHash -Password $SecurePassword -ErrorAction SilentlyContinue
					}
					catch {
						if ($Password.Length -gt 14) {
							Write-Verbose "  ERROR: LMHash couldn't be produced for '$($Password)' as it's >14 characters"
						} else {
							Write-Verbose "  ERROR: LMHash couldn't be produced for '$($Password)'"
						}
						$LMHashError = $TRUE
					}
				}

				# If both hash function produced an error then doesn't produce an output
				if (!$NTHashError -or !$LMHashError) {
					$Query = "INSERT INTO Passwords (Password, PasswordLength, LMHash, NTHash, LowerCase, UpperCase, Digits, Specials) VALUES ('$($SafePassword)', '$($SafePassword.Length)', '$($LMHashCode)', '$($NTHashCode)', '$([int]($SafePassword -cmatch $pfLowerCase))', '$([int]($SafePassword -cmatch $pfUpperCase))', '$([int]($SafePassword -cmatch $pfDigits))', '$([int]($SafePassword -cmatch $pfSpecials))')"

					try {
						Invoke-SqliteQuery -DataSource $SQLiteDB -Query "$($Query)" -ErrorAction SilentlyContinue

						# Update the progress variable
						$PasswordsAdded++
					}
					catch {
						throw "ERROR: Unable to add record for '$($Password)' / '$($SafePassword)'"
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

#	"`nPasswords processed: {0:n0}`n" -f $PasswordsProgressed
	"`nProcessed: {0}" -f $InputFileSizeStr
	"    Added: {0:n0} of {1:n0}`n" -f $PasswordsAdded, $PasswordsProgressed

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
			Comment
			Home Dir

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
        Find-Passwords -InputFile "sample-sam.txt" -SQLiteDatabase "Passwords.SQLite" -ShowProgressBar

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party module isn't possible
#>

	Param (
		[string]$InputFile = $(throw "-InputFile is required."),
		[string]$SQLiteDatabase = $(throw "-SQLiteDatabase is required."),
		[switch]$ShowProgressBar = $false
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
		$FoundLMHash = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM Passwords WHERE LMHash='$($LineParts[2])'"
		$FoundNTHash = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM Passwords WHERE NTHash='$($LineParts[3])'"

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

	"`nPasswords processed: {0:n0}`n" -f $PasswordsProgressed

	# When did the task finish?
	$Finished = Get-Date

	# How long did the work take?
	$HowLong = $Finished - $Started
	"Finished: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Finished.Year, $Finished.Month, $Finished.Day, $Finished.Hour, $Finished.Minute, $Finished.Second
	"Duration: {0:d2}d {1:d2}h {2:d2}m {3:d2}s`n" -f $HowLong.Days, $HowLong.Hours, $HowLong.Minutes, $HowLong.Seconds
}

function Add-MissingData {
<#
    .SYNOPSIS
        This function will create missing hashes in the SQL database

    .DESCRIPTION
		This function will create any missing hashes in the SQL database.
		
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

	.PARAMETER NTHash
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

	.PARAMETER LMHash
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

	.PARAMETER PasswordLength
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

	.PARAMETER LowerCase
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

	.PARAMETER UpperCase
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

	.PARAMETER Digits
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

	.PARAMETER Specials
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
        Create-MissingHashes -SQLiteDatabase "Passwords.SQLite" -ShowProgressBar

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party module isn't possible
#>

	Param (
		[string]$SQLiteDatabase  = $(throw "-SQLiteDatabase is required."),
		[switch]$ShowProgressBar = $false,
		[switch]$NTHash          = $false,
		[switch]$LMHash          = $false,
		[switch]$PasswordLength  = $false,
		[switch]$LowerCase       = $false,
		[switch]$UpperCase       = $false,
		[switch]$Digits          = $false,
		[switch]$Specials        = $false
	)

	# At least 1 switch needs to be $TRUE
	if (!$NTHash -and !$LMHash -and !$PasswordLength -and !$LowerCase -and !$UpperCase -and !$Digits -and !$Specials) {
		# If you get here then all switches are $FALSE
		"INFO: At least one of the following switch parameters needs to be $TRUE"
		"`n`t* NTHash`n`t* LMHash`n`t* PasswordLength`n`t* LowerCase`n`t* UpperCase`n`t* Digits`n`t* Specials`n"

		Break
	}

	[string]$StatusMessage = 'Adding Missing Data for @@MARKER@@'
	if ($NTHash)         {$StatusMessage = $StatusMessage.replace('@@MARKER@@', "NTHash, @@MARKER@@")}
	if ($LMHash)         {$StatusMessage = $StatusMessage.replace('@@MARKER@@', "LMHash, @@MARKER@@")}
	if ($PasswordLength) {$StatusMessage = $StatusMessage.replace('@@MARKER@@', "PasswordLength, @@MARKER@@")}
	if ($LowerCase)      {$StatusMessage = $StatusMessage.replace('@@MARKER@@', "LowerCase, @@MARKER@@")}
	if ($UpperCase)      {$StatusMessage = $StatusMessage.replace('@@MARKER@@', "UpperCase, @@MARKER@@")}
	if ($Digits)         {$StatusMessage = $StatusMessage.replace('@@MARKER@@', "Digits, @@MARKER@@")}
	if ($Specials)       {$StatusMessage = $StatusMessage.replace('@@MARKER@@', "Specials, @@MARKER@@")}
	$StatusMessage = $StatusMessage.replace(', @@MARKER@@', '')

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

	# When did the task start?
	$Started = Get-Date
	"Started: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Started.Year, $Started.Month, $Started.Day, $Started.Hour, $Started.Minute, $Started.Second

	# Used by the Status Bar
	$PasswordsProgressed = 0
	$RecordsUpdated = 0

	# Creste the SQL Search command
	$SearchQuery = 'SELECT * FROM Passwords WHERE @@MARKER@@'
	if ($NTHash)         {$SearchQuery = $SearchQuery.replace('@@MARKER@@', "(NTHash='') OR (NTHash IS NULL) OR @@MARKER@@")}
	if ($LMHash)         {$SearchQuery = $SearchQuery.replace('@@MARKER@@', "(LMHash='' AND PasswordLength < 15) OR (LMHash IS NULL AND PasswordLength < 15) OR @@MARKER@@")}
	if ($PasswordLength) {$SearchQuery = $SearchQuery.replace('@@MARKER@@', "(PasswordLength IS NULL) OR @@MARKER@@")}
	if ($LowerCase)      {$SearchQuery = $SearchQuery.replace('@@MARKER@@', "(LowerCase IS NULL) OR @@MARKER@@")}
	if ($UpperCase)      {$SearchQuery = $SearchQuery.replace('@@MARKER@@', "(UpperCase IS NULL) OR @@MARKER@@")}
	if ($Digits)         {$SearchQuery = $SearchQuery.replace('@@MARKER@@', "(Digits IS NULL) OR @@MARKER@@")}
	if ($Specials)       {$SearchQuery = $SearchQuery.replace('@@MARKER@@', "(Specials IS NULL) OR @@MARKER@@")}
	$SearchQuery = $SearchQuery.replace('OR @@MARKER@@', "LIMIT 1")

	# Work through the database for blanks
	$SelectedRecord = Invoke-SqliteQuery -DataSource $SQLiteDB -Query $SearchQuery -ErrorAction SilentlyContinue
	
	# Is there anything to do?
	while ($SelectedRecord) {
		# Update the progress variables
		$PasswordsProgressed++	

		# Show the progress bar if required
		if ($ShowProgressBar) {
			$StatusText = "Processed: {0:n0} -- {1}" -f $PasswordsProgressed, $($SelectedRecord.Password)
			Write-Progress -Activity $StatusMessage -Status $StatusText
		}

		# Build the SQL command
		$UpdateQuery = 'UPDATE Passwords SET @@MARKER@@,'

		# Update the NTHASH
		if ($NTHash) {
			$NTHashCode = ''

			if (!$SelectedRecord.NTHash -and $SelectedRecord.Password.Length -gt 0) {
				$SecurePassword = ConvertTo-SecureString -String $SelectedRecord.Password -AsPlainText -Force

				try   {
					$NTHashCode = ConvertTo-NTHash -Password $SecurePassword -ErrorAction SilentlyContinue
					$UpdateQuery = $UpdateQuery.replace('@@MARKER@@,', "NTHash='$($NTHashCode)',@@MARKER@@,")
				}
				catch {}
			}
		}

		# Update the LMHASH
		if ($LMHash) {
			$LMHashCode = ''

			if (!$SelectedRecord.LMHash -and $SelectedRecord.Password.Length -gt 0) {
				$SecurePassword = ConvertTo-SecureString -String $SelectedRecord.Password -AsPlainText -Force

				try   {
					$LMHashCode = ConvertTo-LMHash -Password $SecurePassword -ErrorAction SilentlyContinue

					if (!$LMHashCode) {
						# The LMHash generation failed to work which most likely is due to non-latin characters
						# so populate the field with the LM Hash empty hash
						$LMHashCode = $pfEmptyLMHash
					}

					$UpdateQuery = $UpdateQuery.replace('@@MARKER@@,', "LMHash='$($LMHashCode)',@@MARKER@@,")
				}
				catch {}
			}
		}

		if ($PasswordLength) {$UpdateQuery = $UpdateQuery.replace('@@MARKER@@,', "PasswordLength='$($SelectedRecord.Password.Length)',@@MARKER@@,")}
		if ($LowerCase)      {$UpdateQuery = $UpdateQuery.replace('@@MARKER@@,', "LowerCase='$([int]($SelectedRecord.Password -cmatch $pfLowerCase))',@@MARKER@@,")}
		if ($UpperCase)      {$UpdateQuery = $UpdateQuery.replace('@@MARKER@@,', "UpperCase='$([int]($SelectedRecord.Password -cmatch $pfUpperCase))',@@MARKER@@,")}
		if ($Digits)         {$UpdateQuery = $UpdateQuery = $UpdateQuery.replace('@@MARKER@@,', "Digits='$([int]($SelectedRecord.Password -cmatch $pfDigits))',@@MARKER@@,")}
		if ($Specials)       {$UpdateQuery = $UpdateQuery.replace('@@MARKER@@,', "Specials='$([int]($SelectedRecord.Password -cmatch $pfSpecials))',@@MARKER@@,")}
		$UpdateQuery = $UpdateQuery.replace(',@@MARKER@@,', " WHERE ID=$($SelectedRecord.ID)")

		try {
			Invoke-SqliteQuery -DataSource $SQLiteDB -Query $UpdateQuery -ErrorAction SilentlyContinue

			# Update the progress variable
			$RecordsUpdated++
		}
		catch {}

		# Read the next record to process
		try {$SelectedRecord = Invoke-SqliteQuery -DataSource $SQLiteDB -Query $SearchQuery -ErrorAction SilentlyContinue}
		catch {}
	}

	"`nPasswords processed: {0:n0}`n" -f $PasswordsProgressed
	"Records updated:     {0:n0}" -f $RecordsUpdated

	# When did the task finish?
	$Finished = Get-Date

	# How long did the work take?
	$HowLong = $Finished - $Started
	"Finished: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Finished.Year, $Finished.Month, $Finished.Day, $Finished.Hour, $Finished.Minute, $Finished.Second
	"Duration: {0:d2}d {1:d2}h {2:d2}m {3:d2}s`n" -f $HowLong.Days, $HowLong.Hours, $HowLong.Minutes, $HowLong.Seconds
}

function Add-MissingHashes {
<#
    .SYNOPSIS
        This function will create missing hashes in the SQL database

    .DESCRIPTION
		This function will create any missing hashes in the SQL database.
		
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
        Create-MissingHashes -SQLiteDatabase "Passwords.SQLite" -ShowProgressBar

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party module isn't possible
#>

	Param (
		[string]$SQLiteDatabase = $(throw "-SQLiteDatabase is required."),
		[switch]$ShowProgressBar = $false
	)

	Add-MissingData -SQLiteDatabase $SQLiteDatabase -ShowProgressBar $ShowProgressBar -NTHash -LMHash
}

function Add-MissingPasswordLength {
<#
	.SYNOPSIS
		This function will insert the missing PasswordLength in the SQL database

	.DESCRIPTION
		This function will insert the missing PasswordLength in the SQL database.
		
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
		Add-MissingPasswordLength -SQLiteDatabase "Passwords.SQLite" -ShowProgressBar

	.LINK
		Links to further documentation isn't enabled.

	.NOTES
		Error trapping from the 3rd party module isn't possible
#>

	Param (
		[string]$SQLiteDatabase = $(throw "-SQLiteDatabase is required."),
		[switch]$ShowProgressBar = $false
	)

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

	# When did the task start?
	$Started = Get-Date
	"Started: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Started.Year, $Started.Month, $Started.Day, $Started.Hour, $Started.Minute, $Started.Second
	
	# Used by the Status Bar
	$PasswordsProgressed = 0
	$RecordsUpdated = 0

	# Work through the database for blank hashes
	$SelectedRecord = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM Passwords WHERE PasswordLength IS NULL LIMIT 1" -ErrorAction SilentlyContinue

	# Is there anything to do?
	while ($SelectedRecord) {
		# Update the progress variables
		$PasswordsProgressed++	

		# Show the progress bar if required
		if ($ShowProgressBar) {
			$StatusText = "Processed: {0:n0} -- {1}" -f $PasswordsProgressed, $($SelectedRecord.Password)
			Write-Progress -Activity "Setting Password length" -Status $StatusText
		}

		# Update the Record
		$Query = "UPDATE Passwords SET PasswordLength=$($SelectedRecord.Password.Length) WHERE ID=$($SelectedRecord.ID)"
		try {
			Invoke-SqliteQuery -DataSource $SQLiteDB -Query $Query -ErrorAction SilentlyContinue

			# Update the progress variable
			$RecordsUpdated++
		}
		catch {}

		# Read the next record to process
		try {$SelectedRecord = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM Passwords WHERE PasswordLength IS NULL LIMIT 1" -ErrorAction SilentlyContinue}
		catch {}
	}

	"`nPasswords processed: {0:n0}" -f $PasswordsProgressed
	"Records updated:     {0:n0}" -f $RecordsUpdated

	# When did the task finish?
	$Finished = Get-Date

	# How long did the work take?
	$HowLong = $Finished - $Started
	"`nFinished: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Finished.Year, $Finished.Month, $Finished.Day, $Finished.Hour, $Finished.Minute, $Finished.Second
	"Duration: {0:d2}d {1:d2}h {2:d2}m {3:d2}s`n" -f $HowLong.Days, $HowLong.Hours, $HowLong.Minutes, $HowLong.Seconds
}
	
function Find-ExcelPassword {
<#
    .SYNOPSIS
        This function attempts to open an Excel file using the passwords from the Database

    .DESCRIPTION
        This function attempts to open an Excel file using the passwords from the Database
		
		It is possible to select the minimum password length to start on.  This parameter
		was added so time would not be wasted checking passwords that are too short.
		
		This script depends on:
		  DSInternals from https://github.com/MichaelGrafnetter/DSInternals
	      PSSQLite    from https://github.com/RamblingCookieMonster/PSSQLite

    .PARAMETER PasswordFile
        This file containing the list of passowrds.  Each line is considered a password

    .PARAMETER SQLiteDatabase
        This is the SQLite Database to be used

    .PARAMETER MinimumPasswordLength
        This is minimum password length that will be used.

		The default is 1 character as the minimum length

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
        Import-Passwords -InputFile "passwords.txt" -SQLiteDatabase "Passwords.SQLite" -ShowProgressBar

    .LINK
        Links to further documentation isn't enabled.

    .NOTES
		Error trapping from the 3rd party module isn't possible
#>

	Param (
		[string]$ExcelFile = $(throw "-ExcelFile is required."),
		[string]$SQLiteDatabase = $(throw "-SQLiteDatabase is required."),
		[int]$MinimumPasswordLength = 1,
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
	$Password = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT ID,Password FROM Passwords WHERE ID=1 AND PasswordLength >= $MinimumPasswordLength"

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
						$Password = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT ID,Password FROM Passwords WHERE ID = $($Password.ID + 1) AND PasswordLength >= $MinimumPasswordLength"
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

function Import-COMBPasswords {
<#
	.SYNOPSIS
		This function imports new passwords from a text file (: seperated fields) and rejects duplicates

	.DESCRIPTION
		This function imports new passwords from a text file (: seperated fields) and rejects duplicates.
		
		The NTHash and LMHash are NOT created as part of this process.  See 'Add-MissingHashes' for that function.

		This script depends on:
			DSInternals from https://github.com/MichaelGrafnetter/DSInternals
			PSSQLite    from https://github.com/RamblingCookieMonster/PSSQLite

	.PARAMETER InputFile
		This file containing the list of passowrds.

	.PARAMETER SQLiteDatabase
		This is the SQLite Database to be used

	.PARAMETER ShowProgressBar
		This switch is either $TRUE or $FALSE.

		$TRUE  will show a progress bar
		$FALSE will not show a progress bar

		Default value:	$FALSE

    .PARAMETER NTHASH
		This switch is either $TRUE or $FALSE.

		$TRUE  will produce an HASH value
		$FALSE will NOT produce an HASH value

		Default value:	$FALSE

    .PARAMETER LMHASH
		This switch is either $TRUE or $FALSE.

		$TRUE  will produce an HASH value
		$FALSE will NOT produce an HASH value

		Default value:	$FALSE

		Notes:
		The LMHASH function doesn't support Unicode characters and will produce an error

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
		Import-Passwords -InputFile "passwords.txt" -SQLiteDatabase "Passwords.SQLite" -ShowProgressBar

	.LINK
		Links to further documentation isn't enabled.

	.NOTES
		Error trapping from the 3rd party module isn't possible
#>

	Param (
		[string]$InputFile = $(throw "-InputFile is required."),
		[string]$SQLiteDatabase = $(throw "-SQLiteDatabase is required."),
		[switch]$Verbose = $false,
		[switch]$LMHash = $false,
		[switch]$NTHash = $false,
		[switch]$ShowProgressBar = $false
	)

	# Record old Verbose setting
	if ($Verbose) {
		$OldVerbose = $VerbosePreference
		$VerbosePreference = "Continue"
	}

	# When did the task start?
	$Started = Get-Date
	"Started: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Started.Year, $Started.Month, $Started.Day, $Started.Hour, $Started.Minute, $Started.Second

	# How large is the InputFile
	$InputFileSize = (Get-ChildItem $InputFile).Length
	if ($InputFileSize -lt 1GB) {$InputFileSizeStr = "{0:n} MB" -f ($InputFileSize/1MB)} else {$InputFileSizeStr = "{0:n} GB" -f ($InputFileSize/1GB)}
	"Filesize for '$($InputFile)': {0:n}" -f $InputFileSizeStr

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

	# Does the Database need to be created?
	if (!(Test-Path $SQLiteDatabase)) {
		# Remove the Database if it already exists and hide the output
		if (Test-Path $SQLiteDatabase) {
			"Removed: $($SQLiteDatabase)"
			Remove-Item -Path $SQLiteDatabase -Force | Out-Null
		}

		Set-SQLiteDatabase -SQLiteDatabase $SQLiteDB
	} else {
		"Using existing SQLite Database: $($SQLiteDatabase)"
	}

	# Make sure the Database exists
	if (!(Test-Path $SQLiteDatabase)) {
		"ERROR: Missing Database $($SQLiteDatabase)"
		return
	}

	# Used by the Status Bar
	$PasswordsProgressed = 0
	$PasswordsAdded = 0
	$BytesProcessed = 0

	# Work through each line of the folder file removing non ISO-8859-1 characters
	$InputFileWithPath = (Get-ChildItem $InputFile).FullName
	ForEach ($Line in [System.IO.File]::ReadLines($InputFileWithPath))
	{
		# Update the progress variables
		$PasswordsProgressed++

		# Split the line and grab the password from the last block
		[array]$SplitValues = $Line.Split(':')
		$Password = $SplitValues[$SplitValues.Count - 1]

		# Remove non Latin characters
		$Password = $Password -replace '[^\p{IsBasicLatin}\p{IsLatin-1Supplement}]'

		# Reset temporary variables
		$SkipEntry = $FALSE
		$NTHashError = $FALSE
		$LMHashError = $FALSE
		$NTHashCode  = ""
		$LMHashCode  = ""

		# Show the progress bar if required
		if ($ShowProgressBar) {
			$BytesProcessed += $Password.Length

			$PercentageCompleted = ($BytesProcessed/$InputFileSize * 100)
			if ($BytesProcessed -lt 1GB) {$BytesProcessedStr = "{0:n} MB" -f ($BytesProcessed/1MB)} else {$BytesProcessedStr = "{0:n} GB" -f ($BytesProcessed/1GB)}
			$StatusText = "Processed {0:n}% ({1} of {2}) -- Added {3:n0} of {4:n0} -- {5}" -f $PercentageCompleted, $BytesProcessedStr, $InputFileSizeStr, $PasswordsAdded, $PasswordsProgressed, $($Password)
			Write-Progress -PercentComplete $PercentageCompleted -Activity "Processing passwords from $($InputFile)" -Status $StatusText
		}

		# Change the password to stop special characters being processing by SQL
		$SafePassword = Get-SQLSafeText -Text $Password

		# Is the password already in the Database?
		$SelectedRecord = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM Passwords WHERE Password='$($SafePassword)'" -ErrorAction SilentlyContinue

		if ($SelectedRecord) {
			# Record found
			Write-Verbose "  Exists: $($Password) / $($SafePassword)"

			# Remove the temporary variable
			Remove-Variable -Name SelectedRecord
		} else {
			# Encrypt the clear text password for the Hash functions
			try {
				$SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

				# If the password is outside the 0-14 character range it can throw an error so
				# we shall hide those errors
				if ($NTHash) {
					try {
						$NTHashCode = ConvertTo-NTHash -Password $SecurePassword -ErrorAction SilentlyContinue
					}
					catch {
						Write-Verbose "  ERROR: NTHash couldn't be produced for '$($Password)'"
						$NTHashError = $TRUE
					}
				}
				if ($LMHash) {
					try {
						$LMHashCode = ConvertTo-LMHash -Password $SecurePassword -ErrorAction SilentlyContinue
					}
					catch {
						if ($Password.Length -gt 14) {
							Write-Verbose "  ERROR: LMHash couldn't be produced for '$($Password)' as it's >14 characters"
						} else {
							Write-Verbose "  ERROR: LMHash couldn't be produced for '$($Password)'"
						}
						$LMHashError = $TRUE
					}
				}

				# If both hash function produced an error then just record the password
				if ($NTHashError -and $LMHashError) {
					$Query = "INSERT INTO Passwords (Password, PasswordLength, LowerCase, UpperCase, Digits, Specials) VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}')" -f $SafePassword,$SafePassword.Length,$([int]($SelectedRecord.Password -cmatch $pfLowerCase)),$([int]($SelectedRecord.Password -cmatch $pfUpperCase)),$([int]($SelectedRecord.Password -cmatch $pfDigits)),$([int]($SelectedRecord.Password -cmatch $pfSpecials))
				} else {
					$Query = "INSERT INTO Passwords (Password, PasswordLength, LMHash, NTHash, LowerCase, UpperCase, Digits, Specials) VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}')" -f $SafePassword,$SafePassword.Length,$LMHashCode,$NTHashCode,$([int]($SelectedRecord.Password -cmatch $pfLowerCase)),$([int]($SelectedRecord.Password -cmatch $pfUpperCase)),$([int]($SelectedRecord.Password -cmatch $pfDigits)),$([int]($SelectedRecord.Password -cmatch $pfSpecials))
				}

				# Write the record to the database
				try {
					Invoke-SqliteQuery -DataSource $SQLiteDB -Query "$($Query)" -ErrorAction SilentlyContinue

					# Update the progress variable
					$PasswordsAdded++
				}
				catch {
					throw "`nERROR: Unable to add record for '$($Password)' / '$($SafePassword)'"
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

	"`nProcessed: {0}" -f $InputFileSizeStr
	"    Added: {0:n0} of {1:n0}`n" -f $PasswordsAdded, $PasswordsProgressed

	# When did the task finish?
	$Finished = Get-Date

	# How long did the work take?
	$HowLong = $Finished - $Started
	"Finished: {0:d4}/{1:d2}/{2:d2} @ {3:d2}:{4:d2}:{5:d2}" -f $Finished.Year, $Finished.Month, $Finished.Day, $Finished.Hour, $Finished.Minute, $Finished.Second
	"Duration: {0:d2}d {1:d2}h {2:d2}m {3:d2}s`n" -f $HowLong.Days, $HowLong.Hours, $HowLong.Minutes, $HowLong.Seconds
}
