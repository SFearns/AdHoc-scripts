$PSEmailServer = "mailserver"

$Users = Import-Csv -Path "C:\WorkFolder\new-users.csv"

$Users | ForEach-Object {
    $Jump = $False
    Write-Output ("UserID: " + $_.UserID + " (" + $_.DisplayName + ")")

    $ErrorActionPreference = "SilentlyContinue"
    if (Get-ADUser $_.UserID)
    {
        Write-Output ("    Already Exists")
        $Jump = $True
    }
    $ErrorActionPreference = "Continue"

    If ($Jump -eq $False) {
        New-ADUser -Name $_.DisplayName -DisplayName $_.DisplayName `
        -GivenName $_.FirstName -Initials $_.Initials -Surname $_.Surname `
        -AccountPassword (ConvertTo-SecureString -String $_.Password -AsPlainText -Force) `
        -Description $_.Department -EmailAddress ($_.UserID + "@domain.co.uk") -SamAccountName $_.UserID -UserPrincipalName ($_.UserID + "@local.domain") `
        -Path "OU=TestOU,DC=local,DC=domain" `
        -CannotChangePassword $False -ChangePasswordAtLogon $False `
        -PasswordNeverExpires $False -PasswordNotRequired $False -Enabled $True

        if ($_.AccountExpires -gt 0) {
            Write-Output ("    Expiry: " + $_.AccountExpires)
            Set-ADUser -Identity $_.UserID -AccountExpirationDate ($_.AccountExpires + " 23:59:59")
        }

        Write-Output ("    Groups: vDN_Users, vDN_Shared, vDN_Workspaces, XTProxy")
        Add-ADGroupMember -Identity vDN_Users -Members $_.UserID
        Add-ADGroupMember -Identity vDN_Shared -Members $_.UserID
        Add-ADGroupMember -Identity vDN_Workspaces -Members $_.UserID
        Add-ADGroupMember -Identity XTProxy -Members $_.UserID
    }
}
