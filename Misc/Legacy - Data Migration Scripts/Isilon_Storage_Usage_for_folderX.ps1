$ScriptName="Storage usage for Folder X Users"
$ScriptVer="1.0"
$ScriptDate="26th August 2015"
$ScriptAuthor="Stephen Fearns"

# Count must be greater than 0
$HostName=(Get-ChildItem Env:Computername).Value
$CurrentUser=(Get-ChildItem Env:UserName).Value
$Cluster="Isilon"
$Domain="ADDOMAIN"
$QuotaPathToReportOn="*/Folder_X_Users"
$MaxSyncErrors=[int]9999
$MaxSyncResults=[int]9999
$MaxResults=[int]10
$FreeSpaceWarning=[int]80
$FreeSpaceAlert=[int]90
$FreeSpaceCritical=[int]95
$IsilonFreeSpaceWarning=[int]70
$IsilonFreeSpaceAlert=[int]80
$IsilonFreeSpaceCritical=[int]90
$IsilonRootFreeSpaceWarning=[int]85
$IsilonRootFreeSpaceAlert=[int]90
$IsilonRootFreeSpaceCritical=[int]96
$IsilonVarFreeSpaceWarning=[int]75
$IsilonVarFreeSpaceAlert=[int]80
$IsilonVarFreeSpaceCritical=[int]89
$IsilonVarCrashFreeSpaceWarning=[int]75
$IsilonVarCrashFreeSpaceAlert=[int]80
$IsilonVarCrashFreeSpaceCritical=[int]89
$TodaysDate=Get-Date
$IgnoreForDays=30
$HomeFolder = "C:\WorkFolder"
$ReportFolder = "$HomeFolder\Reports - StorageUsage\"
$ReportFileName=[string](Get-Date -Format yyyyMMdd)+"_"+[string](Get-Date -Format HHmmss)+" - $ScriptName ($Cluster).txt"
$HTMLReportFileName=[string](Get-Date -Format yyyyMMdd)+"_"+[string](Get-Date -Format HHmmss)+" - $ScriptName ($Cluster).htm"
$CSVReportFileName=[string](Get-Date -Format yyyyMMdd)+"_"+[string](Get-Date -Format HHmmss)+" - $ScriptName ($Cluster).csv"
$Report=$ReportFolder+$ReportFileName
$HTMLReport=$ReportFolder+$HTMLReportFileName
$CSVReport=$ReportFolder+$CSVReportFileName
$LoadReport=$false
$ReportTitle="$ScriptName ($Cluster)"
$SendAsHTML=$true
$DomainController='DC'
$CourierON='<font face="Courier New, Courier, monospace">'
$CourierOFF='</font>'

# Live Environment
if (Test-Path -Path "$HomeFolder\LoginCreds\$HostName\$CurrentUser\$($Cluster)_LoginCreds - root.xml") {
    $IsilonCreds=Import-CliXML -Path "$HomeFolder\LoginCreds\$HostName\$CurrentUser\$($Cluster)_LoginCreds - root.xml"
    $ClusterUserID=$IsilonCreds.GetNetworkCredential().UserName
    $ClusterPassword=$IsilonCreds.GetNetworkCredential().Password
}
if (Test-Path -Path "$HomeFolder\LoginCreds\$HostName\$CurrentUser\$($Domain)_LoginCreds.xml")  {
    $LoginCreds=Import-CliXML -Path "$HomeFolder\LoginCreds\$HostName\$CurrentUser\$($Domain)_LoginCreds.xml"
}

#Office365 Mail Settings
if (Test-Path -Path "$HomeFolder\LoginCreds\$HostName\$CurrentUser\Office365_LoginCreds.xml") {
    $OfficeCreds=Import-CliXML -Path "$HomeFolder\LoginCreds\$HostName\$CurrentUser\Office365_LoginCreds.xml"
    $MailConnection="SSL"
    $MailServer="fqdn-2-mailserver"
    $MailFrom=$OfficeCreds.GetNetworkCredential().UserName
    $MailPort=587
}

# eMail details
$MailTo="user@domain.co.uk"
$MailBCC="bccuser@domain.co.uk"
$MailSubject=$ReportTitle+" on "+($TodaysDate).DayOfWeek+" "+($TodaysDate).ToLongDateString()+" at "+($TodaysDate).ToShortTimeString()

# HTML report variables
# Wheat
$CSS='<style>
             table{margin:auto; width:95%; Text-align:Center;}
              Body{background-color:PapayaWhip; Text-align:Center;}
       tr:hover td{background-color:rgb(150, 150, 220); color:black;}
tr:nth-child(even){background-color:rgb(242, 242, 242);}
                th{background-color:DeepSkyBlue; color:black; Text-align:Center;}
                td{background-color:Gainsboro; color:Black; Text-align:Center;}

         table.toc{margin:auto; width:auto; Text-align:Center;}
    table.toc Body{text-align:Center}
      table.toc td{color:Black; text-align:Center;}

          table.ov{margin:auto; width:auto}
       table.ov td{color:Black; padding: 5px;}
</style>'
$ColourBlackOn='<font color=Black>'
$ColourBlackOff='</font>'
$ColourGreenOn='<font color=Green><b>'
$ColourGreenOff='</b></font>'
$ColourWarningOn='<font color=Orange><b>'
$ColourWarningOff='</b></font>'
$ColourAlertOn='<font color=DarkRed><b>'
$ColourAlertOff='</b></font>'
$ColourCriticalOn='<font color=Red><b>'
$ColourCriticalOff='</b></font>'

function Out-HTML {
param([Parameter(Mandatory=$false,ValueFromPipeline=$true)]  [string[]]$Text=$null,
      [Parameter(Mandatory=$false,ValueFromPipeline=$false)] [string]$PreContent=$null,
      [Parameter(Mandatory=$false,ValueFromPipeline=$false)] [string]$Path=$null)
    $Result=$null
    if ($PreContent) {$Result=$PreContent}
    if ($Text) {for ($i=0;$i-lt$Text.Count;$i++) {$Result+=$Text[$i]+'<br>'}}
    if ($Path -and ($PreContent -or $Text)) {$Result | Add-Content -Path $Path;return}
    return $Result
}

# Connect to the Isilon Cluster
$ifsConnection = Connect-IsilonCluster -ClusterName $Cluster -Username $ClusterUserID -Password $ClusterPassword

if ($ifsConnected -like "Unable to connect to*") {return $ifsConnected}
if ($ifsConnected -like "No SSH session found*") {return $ifsConnected}

if ($LoginCreds -and $DomainController) {
	Write-Output "Gathering AD information"
	try {$ADUsers=Get-ADUser -Filter * -Properties * -Server $DomainController -Credential $LoginCreds -ErrorAction SilentlyContinue}
	catch {Write-Host "Not able to connect to $DomainController"}
}

ConvertTo-Html -Title "$ReportTitle" -Head "<h1>$ReportTitle<br></h1><br>This report was created at $(Get-Date)<br>$ScriptName v$ScriptVer by $ScriptAuthor" -Body "$CSS" | Set-Content -Path $HTMLReport

$TempQuotaList=Get-IsilonListQuotas -ClusterName $Cluster | Sort-Object Path

# Produce the CSV file
$TempQuotaList | Where-Object {($PSItem.type -eq 'user') -and ($PSItem.path -like $QuotaPathToReportOn)} | Select appliesto,path,usage_derived | Export-Csv -Path $CSVReport
if (!(Test-Path -Path $CSVReport)) {"Error producing CSV file" | Add-Content -Path $CSVReport}

# "<h2>User quotas</h2>" | Add-Content -Path $HTMLReport
"<table>" | Add-Content -Path $HTMLReport
"<colgroup><col/><col/><col/><col/><col/></colgroup>" | Add-Content -Path $HTMLReport
"<tr><th>User</th><th>Path</th><th>Used</th></tr>" | Add-Content -Path $HTMLReport
$TempQuotaList | Where-Object {($PSItem.type -eq 'user') -and ($PSItem.path -like $QuotaPathToReportOn)} | Sort-Object appliesto,path | ForEach-Object{
    $Size="TB"; $Usage=($PSItem.usage_derived / 1TB)
    if ($PSItem.usage_derived -le (1TB-1GB)){$Size="GB"; $Usage=($PSItem.usage_derived / 1GB)}
    if ($PSItem.usage_derived -le (1GB-1MB)){$Size="MB"; $Usage=($PSItem.usage_derived / 1MB)}
    $Path=$PSItem.path
    $User=$PSItem.appliesto
    Write-Progress -Activity "Searching for Disabled AD accounts" -CurrentOperation $User
    if ($ADUsers) {
        $ADUserID=$ADUsers | Where-Object SamAccountName -eq $User.Replace("$Domain\",'')
        if ($?) {
            if (($ADUserID.AccountExpirationDate -and 
                ($ADUserID.AccountExpirationDate -lt $TodaysDate.AddDays(0-$IgnoreForDays))) -or 
                !$ADUserID.Enabled) {
                $c1=$ColourCriticalOn
                $c2=$ColourCriticalOff
            } else {
                $c1=$null
                $c2=$null
            }
        }
    } else {
        $c1=$null
        $c2=$null
    }
    "<tr><td>{4}{0}{5}</td><td>{1}</td><td>{2:N2} {3}</td></tr>" -f $User,$Path,$Usage,$Size,$c1,$c2 | Add-Content -Path $HTMLReport
}
"</table>" | Add-Content -Path $HTMLReport

$ifsConnection = Disconnect-IsilonCluster -ClusterName $Cluster

if ($MailServer -and $MailFrom -and $MailTo){
    if ($SendAsHTML) {
        $t=(Get-Content $HTMLReport)
		if ($MailConnection -like 'SSL') {
			if ($MailBCC){
				Send-MailMessage -Credential $OfficeCreds -useSSL -SmtpServer $MailServer -Port $MailPort -From $MailFrom -To $MailTo -BCC $MailBCC -Subject $MailSubject -BodyAsHtml "$t<br><br>" -Attachments $HTMLReport,$CSVReport
			} else {
				Send-MailMessage -Credential $OfficeCreds -useSSL -SmtpServer $MailServer -Port $MailPort -From $MailFrom -To $MailTo -Subject $MailSubject -BodyAsHtml "$t<br><br>" -Attachments $HTMLReport,$CSVReport
			}
		} else {
			if ($MailBCC){
				Send-MailMessage -SmtpServer $MailServer -Port $MailPort -From $MailFrom -To $MailTo -BCC $MailBCC -Subject $MailSubject -BodyAsHtml "$t<br><br>" -Attachments $HTMLReport,$CSVReport
			} else {
				Send-MailMessage -SmtpServer $MailServer -Port $MailPort -From $MailFrom -To $MailTo -Subject $MailSubject -BodyAsHtml "$t<br><br>" -Attachments $HTMLReport,$CSVReport
			}

		}
    } else {
		if ($MailConnection -like 'SSL') {
			if ($MailBCC){
				Send-MailMessage -Credential $OfficeCreds -useSSL -SmtpServer $MailServer -Port $MailPort -From $MailFrom -To $MailTo -BCC $MailBCC -Subject $MailSubject -Body "Please find attached the report.`n" -Attachments $HTMLReport,$CSVReport
			} else {
				Send-MailMessage -Credential $OfficeCreds -useSSL -SmtpServer $MailServer -Port $MailPort -From $MailFrom -To $MailTo -Subject $MailSubject -Body "Please find attached the report.`n" -Attachments $HTMLReport,$CSVReport
			}
		} else {
			if ($MailBCC){
				Send-MailMessage -SmtpServer $MailServer -Port $MailPort -From $MailFrom -To $MailTo -BCC $MailBCC -Subject $MailSubject -Body "Please find attached the report.`n" -Attachments $HTMLReport,$CSVReport
			} else {
				Send-MailMessage -SmtpServer $MailServer -Port $MailPort -From $MailFrom -To $MailTo -Subject $MailSubject -Body "Please find attached the report.`n" -Attachments $HTMLReport,$CSVReport
			}
		}
    }
}

If ($LoadReport-eq$true) {Invoke-Item -Path $HTMLReport}