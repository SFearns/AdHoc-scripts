##################################################################
## Copyright 2024  Stephen Fearns
##################################################################

# Load the Required modules
Import-Module PSSQLite

# Variable defination block
$SQLiteDB = ".\Database.SQLite"
$HostName = HostName.EXE

# Make sure the Database exists
if (!(Test-Path $SQLiteDB)) {
    "ERROR: Missing Database $($SQLiteDB)"
    Break;
}

# Load the commands for this host (Server=$Hostname) and commands for everyone (Server='')
$WorkList = Invoke-SqliteQuery -DataSource $SQLiteDB -Query "SELECT * FROM Commands WHERE (Server='$($HostName)' OR Server='')" -ErrorAction SilentlyContinue

# Were there any commands to execute?
if ($WorkList) {
    "Executing $($WorkList.count) commands"
    $WorkList | Foreach-Object -ThrottleLimit 5 -Parallel {
        #Action that will run in Parallel. Reference the current object via $PSItem and bring in outside variables with $USING:varname

        # Build up the Command and remove leading / trailing spaces
        if ($PSItem.Parameters) {
            $Command = ($PSItem.Command + " " + $PSItem.Parameters).Trim()
        } else {
            $Command = ($PSItem.Command).Trim()
        }

        # Execute the command recording how long it took
        "Running: $($Command)"
        $ExecutionTime = Measure-Command {$Result = Invoke-Expression $Command}
        
        # Build the data for the Record
        $Temp = Get-Date
        $DateTime = "{0} {1}" -f $Temp.ToShortDateString(),$Temp.ToShortTimeString()
        $Server = $USING:Hostname
        $JSON = $Result | ConvertTo-Json
        $RAW = $Result | Out-String

        # Store the result of the command in the SQLite Database
        # $Query = "INSERT INTO MachineData (DateTime, Server, Command, ExecutionTime, JSON, RAW) VALUES ('{0}','{1}','{2}','{3},'{4},'{5}')" -f $DateTime,$Server,$Command,$ExecutionTime,$JSON.Replace('"','\"'),$RAW
        $Query = "INSERT INTO MachineData (DateTime, Server, Command, ExecutionTime, JSON, RAW) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}')" -f $DateTime,$Server,$Command,$ExecutionTime,$JSON,$RAW

        try {Invoke-SqliteQuery -DataSource $USING:SQLiteDB -Query $Query} # -ErrorAction SilentlyContinue}
        catch {"ERROR: Unable to write data in $($USING:SQLiteDB)"}
    }
}

# End of file