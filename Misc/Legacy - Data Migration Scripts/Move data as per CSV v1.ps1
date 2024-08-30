Import-Module NTFSSecurity
Import-Module IsilonFunctions

$HomeFolder='C:\WorkFolder'
$HostName=(Get-ChildItem Env:Computername).Value
$CurrentUser=(Get-ChildItem Env:UserName).Value
$Domain='ADDOMAIN'
$MyAdminAccount='Administrator'
$IsilonCluster='Isilon'
$IsilonRootSMBShare='SMBShare'
$IsilonMountPoint="\\$IsilonCluster\$IsilonRootSMBShare"
$IsilonRootPath='RootFolder'
$IsilonSourceZone ='SOURCE1'
$IsilonSourceZone2='SOURCE2'
$IsilonTargetZone ='TARGET1'
$FolderREADME="$HomeFolder\Data Migration Scripts\README for Folder moves.txt"
$FileREADME="$HomeFolder\Data Migration Scripts\README for File moves.txt"

if (Test-Path -Path "$HomeFolder\LoginCreds\$HostName\$CurrentUser\$($IsilonCluster)_LoginCreds - root.xml") {
    $IsilonCreds=Import-CliXML -Path "$HomeFolder\LoginCreds\$HostName\$CurrentUser\$($IsilonCluster)_LoginCreds - root.xml"
    $ClusterUserID=$IsilonCreds.GetNetworkCredential().UserName
    $ClusterPassword=$IsilonCreds.GetNetworkCredential().Password
}
$WorkLoad = [array](Import-Csv -Path "$HomeFolder\Data Migration Scripts\Move_These_Objects.csv") | Sort-Object Type

# Connect to the Isilon Cluster
$ifsConnection = Connect-IsilonCluster -ClusterName $IsilonCluster -Username $ClusterUserID -Password $ClusterPassword
if (($ifsConnection -like "Unable to connect to*") -or ($ifsConnection -like "No SSH session found*")) {
    Write-Output "ERROR: $ifsConnection"
} else {
    Clear-Host

    Write-Output "Gather information about the Isilon Quota settings"
    $QuotaList=Get-IsilonListQuotas -ClusterName $IsilonCluster | Sort-Object Path

    if ($WorkLoad) {
        # $IsilonDrive = New-PSDrive -Name Isilon -PSProvider FileSystem -Root $IsilonMountPoint -Credential $IsilonCreds

        for ($i=0;$i-lt$WorkLoad.Count;$i++) {
            if ($WorkLoad[$i].Type.ToLower().trim() -eq 'file') {
                $WorkLoad[$i] | Format-List
                $Skip=$false
                $SourcePath=$WorkLoad[$i].Path.Replace('/ifs/','').Replace('/','\')
                $SourceFile="$IsilonMountPoint\$SourcePath\"+$WorkLoad[$i].FileName
                $TargetPath="$IsilonMountPoint\"+$WorkLoad[$i].Path.Replace('/ifs/','').Replace("/$IsilonSourceZone/","/$IsilonTargetZone/").Replace("/$IsilonSourceZone2/","/$IsilonTargetZone/").Replace('/','\')

                # Does the source file exist?
                if (!(Test-Path -Path $SourceFile)) {
                    Write-Output "`tMISSING:`t$SourceFile"
                } else {
                    # Does the target folder exist?
                    if (!(Test-Path -path $TargetPath)) {
                        # Break down the path into the individual folders
                        $SourceFolders=$SourcePath.split('\')
                        $sdn="$IsilonMountPoint\$IsilonRootPath\$IsilonSourceZone"
                        $tdn="$IsilonMountPoint\$IsilonRootPath\$IsilonTargetZone"
                        Write-Output "SOURCE: $SourceFile"
                        Write-Output "TARGET: $tdn"
                        for ($i2=2;$i2-lt$SourceFolders.count;$i2++) {
                            $Skip=$false
                            $sdn+="\"+$SourceFolders[$i2]
                            $tdn+="\"+$SourceFolders[$i2]
                            if (Test-Path -path $tdn) {
                                Write-Output "`tFOUND:`t$($SourceFolders[$i2])"
                            } else {
                                Write-Output "`t$($SourceFolders[$i2])"
                                New-Item -Path $tdn -ItemType Directory > $null
                                if (Test-Path -path $tdn) {
                                    $MyAdminAdded=$tacl=$null;$tacl = Get-Item $tdn | Get-Access
                                    if ((Get-Item $sdn).IsInheritanceBlocked) {
                                        Write-Output "`t`tInheritance Blocked"
                                        $tacl | Add-Access -Account "$Domain\$MyAdminAccount" -AccessRights Full
                                        $MyAdminAdded=$true
                                        $tacl | Disable-AccessInheritance
                                    } else {
                                        Write-Output "`t`tInheritance Allowed"
                                    }
                                    Copy-Access -Path $sdn -DestinationPath $tdn
                                } else {
                                    Write-Output "`tERROR:`tNot able to create the target folder"
                                    $Skip=$true; $i2=$SourceFolders.count
                                }
                            }
                        }
                    } else {
                        Write-Output "`tFOUND:`t$TargetPath"
                    }
                    if ($Skip -eq $false) {
                        if (Test-Path -path "$TargetPath\$($WorkLoad[$i].FileName)") {
                            Write-Output "`tERROR:`tFile already exists in the target location - ORIGINAL LEFT UNCHANGED."
                        } else {
                            # Move the file
                            Move-Item -Path $SourceFile -Destination $TargetPath
                            if (Test-Path -path "$TargetPath\$($WorkLoad[$i].FileName)") {
                                Write-Output "`tFile Moved"

                                if (Test-Path -path $SourceFile) {
                                    Write-Output "`tERROR:`tSource File still exists - MOVE failed."
                                } else {
                                    # Copy README file
                                    Copy-Item -Path $FileREADME -Destination "$SourceFile -- README.txt"
                                }
                            } else {
                                Write-Output "`tERROR:`tFile could not be moved."
                            }
                        }
                    }
                }
            }
            if ($WorkLoad[$i].Type.ToLower().trim() -eq 'folder') {
                $WorkLoad[$i] | Format-List
                $Skip=$false
                $SourcePath=$WorkLoad[$i].Path.Replace('/ifs/','').Replace('/','\')
                $TargetPath="$IsilonMountPoint\"+$WorkLoad[$i].Path.Replace('/ifs/','').Replace("/$IsilonSourceZone/","/$IsilonTargetZone/").Replace("/$IsilonSourceZone2/","/$IsilonTargetZone/").Replace('/','\')

                # Does the source folder exist?
                if (!(Test-Path -Path "$IsilonMountPoint\$SourcePath")) {
                    Write-Output "`tMISSING:`t$IsilonMountPoint\$SourcePath"
                } else {
                    # Does the target folder exist?
                    if (!(Test-Path -path $TargetPath)) {
                        # Break down the path into the individual folders
                        $SourceFolders=$SourcePath.split('\')
                        $sdn="$IsilonMountPoint\$IsilonRootPath\$IsilonSourceZone"
                        $tdn="$IsilonMountPoint\$IsilonRootPath\$IsilonTargetZone"
                        Write-Output "SOURCE: $SourcePath"
                        Write-Output "TARGET: $tdn"
                        for ($i2=2;$i2-lt($SourceFolders.count-1);$i2++) {
                            $Skip=$false
                            $sdn+="\"+$SourceFolders[$i2]
                            $tdn+="\"+$SourceFolders[$i2]
                            if (Test-Path -path $tdn) {
                                Write-Output "`tFOUND:`t$($SourceFolders[$i2])"
                            } else {
                                Write-Output "`t$($SourceFolders[$i2])"
                                New-Item -Path $tdn -ItemType Directory > $null
                                if (Test-Path -path $tdn) {
                                    $MyAdminAdded=$tacl=$null;$tacl = Get-Item $tdn | Get-Access
                                    if ((Get-Item $sdn).IsInheritanceBlocked) {
                                        Write-Output "`t`tInheritance Blocked"
                                        $tacl | Add-Access -Account "$Domain\$MyAdminAccount" -AccessRights Full
                                        $MyAdminAdded=$true
                                        $tacl | Disable-AccessInheritance
                                    } else {
                                        Write-Output "`t`tInheritance Allowed"
                                    }
                                    Copy-Access -Path $sdn -DestinationPath $tdn
                                } else {
                                    Write-Output "`tERROR:`tNot able to create the target folder"
                                    $Skip=$true; $i2=$SourceFolders.count
                                }
                            }
                        }
                        $TargetPath=$tdn
                    } else {
                        Write-Output "`tFOUND:`t$TargetPath"
                        $Skip=$true
                    }
                    if ($Skip -eq $false) {
                        if (Test-Path -path "$TargetPath\$($SourceFolders[$SourceFolders.count-1])") {
                            Write-Output "`tERROR:`tFOLDER already exists in the target location - ORIGINAL LEFT UNCHANGED."
                        } else {
                            # Does a Quota exist on the Isilon?  If so then remove it.
                            [array]$QuotaSubset=$null
                            $TempSourcePath = "$IsilonMountPoint\$SourcePath".Replace('\','/').Replace('//Isilon','')
                            for ($qli=0;$qli-lt$QuotaList.count;$qli++) {
                                # We use 'default-user' so the user quota settings can be ignored
                                if (($QuotaList[$qli].type -ne 'user') -and ($TempSourcePath -like "$($QuotaList[$qli].path)*")) {
                                    # Compile a list of elements which have been removed
                                    [array]$QuotaSubset+=$qli
                                }
                            }
                            # $QuotaSubset is $null unless it exists
                            if ($QuotaSubset) {
                                for ($qli=0;$qli-lt$QuotaSubset.count;$qli++) {
                                    Remove-IsilonQuota -ClusterName $IsilonCluster -Path $QuotaList[$QuotaSubset[$qli]].path -Type $QuotaList[$QuotaSubset[$qli]].type -Detailed $True
                                }
                            }
                            # Move the file
                            Move-Item -Path "$IsilonMountPoint\$SourcePath" -Destination $TargetPath # -Credential $IsilonCreds

                            # if 'mv' command is used then change \ for /
                            # $Result = (([string](Invoke-SshCommand -ComputerName $ClusterName -Command "mv `'$($IsilonMountPoint.replace('\','/'))\$($SourcePath.replace('\','/'))`' `'$($TargetPath.replace('\','/'))`'").split("`r")).Split("`n"))
                            # $Result

                            if (Test-Path -path "$TargetPath\$($SourceFolders[$SourceFolders.count-1])") {
                                Write-Output "`tFolder Moved"
                                # Re-Create the original folder
                                # Copy README file into the re-created folder
                                if (Test-Path -path "$IsilonMountPoint\$SourcePath") {
                                    Write-Output "`tERROR:`tSource Path still exists - MOVE failed."
                                } else {
                                    # Copy README file
                                    New-Item -Path "$IsilonMountPoint\$SourcePath" -ItemType Directory > $null
                                    if (Test-Path -path "$IsilonMountPoint\$SourcePath") {
                                        $MyAdminAdded=$tacl=$null;$tacl = Get-Item "$IsilonMountPoint\$SourcePath" | Get-Access
                                        if ((Get-Item "$TargetPath\$($SourceFolders[$SourceFolders.count-1])").IsInheritanceBlocked) {
                                            Write-Output "`t`tInheritance Blocked"
                                            $tacl | Add-Access -Account "$Domain\$MyAdminAccount" -AccessRights Full
                                            $tacl | Disable-AccessInheritance
                                        } else {
                                            Write-Output "`t`tInheritance Allowed"
                                        }
                                        Copy-Access -Path "$TargetPath\$($SourceFolders[$SourceFolders.count-1])" -DestinationPath "$IsilonMountPoint\$SourcePath"
                                        Copy-Item -Path $FolderREADME -Destination "$IsilonMountPoint\$SourcePath\README.txt"
                                    } else {
                                        Write-Output "`tERROR:`tUnable to re-create source folder."
                                        Copy-Item -Path $FolderREADME -Destination "$IsilonMountPoint\$SourcePath -- README.txt"
                                    }
                                }
                            } else {
                                Write-Output "`tERROR:`tFolder could not be moved."
                            }
                            # Put the Quota back on the Isilon
                            if ($QuotaSubset) {
                                for ($qli=0;$qli-lt$QuotaSubset.count;$qli++) {
                                    # Don't create Quota settings for the user accounts as 'default-user' will do that
                                    if ($QuotaList[$QuotaSubset[$qli]].type -ne 'user') {
                                        New-IsilonQuota -ClusterName $IsilonCluster -Path $QuotaList[$QuotaSubset[$qli]].path -Type $QuotaList[$QuotaSubset[$qli]].type -Container $QuotaList[$QuotaSubset[$qli]].container -Enforced $QuotaList[$QuotaSubset[$qli]].enforced -Snapshots $QuotaList[$QuotaSubset[$qli]].include_snapshots -Overhead $QuotaList[$QuotaSubset[$qli]].thresholds_include_overhead -HardThreshold $QuotaList[$QuotaSubset[$qli]].thresholds.hard -AdviseThreshold $QuotaList[$QuotaSubset[$qli]].thresholds.advisory -SoftThreshold $QuotaList[$QuotaSubset[$qli]].thresholds.soft -SoftGrace $QuotaList[$QuotaSubset[$qli]].thresholds.soft_grace -Detailed $True
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        if ($IsilonDrive) {Remove-PSDrive -Name Isilon}
    }
    $ifsConnection = Disconnect-IsilonCluster -ClusterName $IsilonCluster
}