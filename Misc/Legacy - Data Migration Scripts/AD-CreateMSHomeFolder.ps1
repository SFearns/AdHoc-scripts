<#   
.SYNOPSIS
    Create a private folder
         
.DESCRIPTION
    Create a private folder

.PARAMETER Computername
    Name of the user in the ADDOMAIN domain and the HD Ticket number
         
.NOTES
    Author: Stephen Fearns
    Version: 1.0
        - Initial Script
   
.EXAMPLE
    AD-CreateMSHomeFolder -UserID LoginID -TicketID TicketNumber
#>         

function SF-CreateMSHomeFolder {
    [CmdletBinding ()]
    Param([Parameter(Mandatory =$true)] [string]$UserID,
          [Parameter(Mandatory =$true)] [int]$TicketID  )

    Set-StrictMode -Version 2.0

    Write-Verbose "Create private folder for [$UserID]"
    $UserFolder = '\\TESTDC\MS\' + $UserID
    New-Item -Path $UserFolder -ItemType Directory > $null

    Write-Verbose "Set [$UserID]:Modify permissions on [$UserFolder]"
    $acl = Get-Acl $UserFolder
    $acl. SetAccessRuleProtection($True, $True)
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule ($UserID, 'Modify', 'ContainerInherit, ObjectInherit', 'None', 'Allow')
    $acl. AddAccessRule($rule)
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule ($UserID, 'Delete', 'None, None', 'None', 'Deny')
    $acl. AddAccessRule($rule)
    Set-Acl $UserFolder $acl

    Write-Verbose "Removing AD Group: MS Isilon Group"
    $acl = Get-Acl $UserFolder
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule ('MS Isilon Group', 'ReadAndExecute', 'Allow')
    $acl. RemoveAccessRuleAll($rule)
    Set-Acl $UserFolder $acl

    Write-Verbose "Removing AD Group: MSR_Group_Isilon"
    $acl = Get-Acl $UserFolder
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule ('MSR_Group_Isilon', 'ReadAndExecute', 'Allow')
    $acl. RemoveAccessRuleAll($rule)
    Set-Acl $UserFolder $acl

    Write-Verbose "Removing AD Group: MSR_Group_Isilon"
    $acl = Get-Acl $UserFolder
    $rule = New-Object System.Security.AccessControl.FileSystemAccessRule ('MSR_Group_Isilon_ReadOnly', 'ReadAndExecute', 'Allow')
    $acl. RemoveAccessRuleAll($rule)
    Set-Acl $UserFolder $acl

    Write-Verbose "Send an email to Ticketing System"
    Send-MailMessage -From 'user@domain.co.uk' -To 'helpdesk@ticketing.co.uk' -Subject "[TICK:$TicketID ] Private folder requested" -Body "Private folder ($UserFolder ) created" -Port 25 -SmtpServer 'mailserver'
}