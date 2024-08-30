$OUs = Import-Clixml -Path "C:\WorkFolder\OUs.xml"

$OUs | ForEach-Object {
    $Jump = $False
    Write-Output ("OU: " + $_.DistinguishedName)

    $ErrorActionPreference = "SilentlyContinue"
    if (Get-ADObject $_.DistinguishedName)
    {
        Write-Output ("    Already Exists`n")
        $Jump = $True
    }
    $ErrorActionPreference = "Continue"

    If ($Jump -eq $False) {
        $Protect = $False
        if ($_.ProtectedFromAccidentalDeletion -eq "true") {
            $Protect = $True
        }

        New-ADObject -name $_.Name -Type container -Path "DC=local,DC=domain" -Description $_.Description -DisplayName $_.Name -ProtectedFromAccidentalDeletion $Protect
    }
}
