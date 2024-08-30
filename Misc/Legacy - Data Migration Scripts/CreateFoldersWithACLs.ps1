Import-Module NTFSSecurity

$HomeFolder = "C:\WorkFolder"
$HostName=(Get-ChildItem Env:Computername).Value
$CurrentUser=(Get-ChildItem Env:UserName).Value
$Domain = 'ADDOMAIN'
$MyAdminAccount = 'administrator'

$SourceDir = '\\SourceServer\SMBShare'
$TargetDir = '\\TargetServer\SMBShare\folder1\folder2\folder3\folder4'

if (Test-Path -Path "$HomeFolder\LoginCreds\$HostName\$CurrentUser\$($Domain)_LoginCreds.xml")  {
    $LoginCreds=Import-CliXML -Path "$HomeFolder\LoginCreds\$HostName\$CurrentUser\$($Domain)_LoginCreds.xml"

    $Source = New-PSDrive -Name SourceConnection -PSProvider FileSystem -Root $SourceDir -Credential $LoginCreds # -ErrorAction:SilentlyContinue
    $Target = New-PSDrive -Name TargetConnection -PSProvider FileSystem -Root $TargetDir -Credential $LoginCreds # -ErrorAction:SilentlyContinue

    $SourceDirList = Get-ChildItem -Path $SourceDir -Directory -Force

    for ($i=0;$i-lt$SourceDirList.count;$i++) {
        $sdn=$SourceDir + '\' + $SourceDirList[$i]
        $tdn=$TargetDir + '\' + $SourceDirList[$i].Name
        Write-Output "$sdn"
        Write-Output "`t$tdn"
        if (Test-Path -path $tdn) {
            Write-Output "`tTarget folder already exists"
        } else {
            New-Item -Path $tdn -ItemType Directory > $null
        }

        if (Test-Path -path $tdn) {
            $tacl=$null;$tacl = Get-Item $tdn | Get-Access
            if ($SourceDirList[$i].IsInheritanceBlocked) {
                Write-Output "`tInheritance Blocked"
                $tacl | Add-Access -Account "$Domain\$MyAdminAccount" -AccessRights Full
            } else {
                Write-Output "`tInheritance Allowed"
            }
            $tacl | Disable-AccessInheritance
            Copy-Access -Path $sdn -DestinationPath $tdn
            if ($SourceDirList[$i].IsInheritanceBlocked) {
                $tacl | Remove-Access -Account "$Domain\$MyAdminAccount"
            }
        } else {
            Write-Output "`tERROR: Not able to create the target folder"
        }
    }

    Remove-PSDrive -Name SourceConnection
    Remove-PSDrive -Name TargetConnection
} else {
    Write-Output 'ERROR: Missing the Login Credentials file'
}