<#
.SYNOPSIS
.DESCRIPTION
-Set expiration dates for retired/terminated employee user accounts.
-Sends expiration notifications to users with expiring accounts.
-Disables expired accounts.
-Removes accounts expired for over x months
.EXAMPLE
.INPUTS
.OUTPUTS
.NOTES
Special Thanks to June Castillote for helping me get MSOnline functioning properyl in Jenkins!
https://adamtheautomator.com/jenkins-powershll-64bit/
Vasil Michev
https://www.michev.info/Blog/Post/1655/quickly-list-all-groups-a-user-is-member-of-or-owner-of-in-office-365
#>
[cmdletbinding()]
param (
 [Parameter(Mandatory = $True)][string[]]$AccountManagerEmail,
 [Parameter(Mandatory = $True)][Alias('ADCred')][System.Management.Automation.PSCredential]$ActiveDirectoryCredential,
 [Parameter(Mandatory = $True)][Alias('DCs')][string[]]$DomainControllers,
 [Parameter(Mandatory = $True)][Alias('License')][string[]]$DefaultGroups,
 [Parameter(Mandatory = $True)][Alias('MailCred')][System.Management.Automation.PSCredential]$EmailCredential,
 [Parameter(Mandatory = $True)][System.Management.Automation.PSCredential]$EscapeCredential,
 [Parameter(Mandatory = $True)][string]$EscapeDatabase,
 [Parameter(Mandatory = $True)][string]$EscapeServer,
 [Parameter(Mandatory = $True)][Alias('ExchCred')][System.Management.Automation.PSCredential]$ExchangeOnlineCredentials,
 [Parameter(Mandatory = $True)][int]$GracePeriodDays,
 [Parameter(Mandatory = $True)][int]$StaleSubMonths,
 [Parameter(Mandatory = $True)][int]$DisabledMaxMonths,
 [Parameter(Mandatory = $True)][System.Management.Automation.PSCredential]$SISCredential,
 [Parameter(Mandatory = $True)][string]$SISDatabase,
 [Parameter(Mandatory = $True)][string]$SISServer,
 [Parameter(Mandatory = $True)][string]$StaffOU,
 [string[]]$BccAddress,
 [Alias('wi')][SWITCH]$WhatIf
)

function Add-DefaultGroups ($groups) {
 process {
  foreach ($group in $groups) {
   Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.Name, $_.ad.EmployeeID, $_.ad.Mail, $group) -f Blue
   Add-ADGroupMember -Identity $group -Members $_.ad.ObjectGUID
  }
  $_
 }
}

function Add-EmpData ($dbParams) {
 begin {
  $sql = 'SELECT EmailHome from vwHREmploymentList WHERE EmpID = @id'
 }
 process {
  if ($_.ad.EmployeeID -notmatch '\d') {
   Write-Host ('{0},{1},EMployeeID not an integer [{2}]' -f $MyInvocation.MyCommand.Name, $_.ad.Mail, $_.ad.EmployeeID) -F Red
   return $_
  }
  $sqlVars = "id=$($_.ad.EmployeeID)"
  $empData = New-SqlOperation @dbParams -Query $sql -Parameters $sqlVars | ConvertTo-Csv | ConvertFrom-Csv
  if ($empData.EmailHome -like '*@*') {
   $_.emp += $empData
  }
  $_
 }
}

function Clear-ExtensionAttribute1 {
 process {
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.Mail) -F DarkYellow
  Set-ADUser -Identity $_.ad.ObjectGUID -Clear extensionAttribute1 -WhatIf:$WhatIf
  $_
 }
}

function Disable-ADAccount {
 process {
  $msgData = $MyInvocation.MyCommand.name, $_.ad.EmployeeID , $_.ad.Mail, $_.ad.AccountExpirationDate
  Write-Host ('{0},{1},{2},AccountExpirationDate [{3}]' -f $msgData ) -F DarkYellow
  Set-ADUser -Identity $_.ad.ObjectGUID -Enabled:$false -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Disable-PreventAccidentalDelete {
 process {
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.ad.Mail) -F Yellow
  $prevParams = @{
   Identity                        = $_.ad.ObjectGUID
   ProtectedFromAccidentalDeletion = $False
   Confirm                         = $False
   ErrorAction                     = 'Stop'
   WhatIf                          = $WhatIf
  }
  Set-ADObject @prevParams
  $_
 }
}

function Disable-GSuiteForwarding {
 begin {
  $gam = '.\bin\gam.exe'
 }
 process {
  # TODO scope - access denied issue - gamadv might work instead
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.ad.EmployeeID , $_.homepage ) -F DarkYellow
  Write-Host ("& $gam user {0} forward off" -f $_.homepage) -F Yello
  if (!$WhatIf) { & $gam user $_.homepage forward off }
  $_
 }
}

function Format-EmailExpiring {
 begin {
  $baseMsg = Get-Content -Path .\html\EmailExpiring.html -Raw
  $head = "<style TYPE=`"TEXT/CSS`"><!--`n" + (Get-Content -Path .\html\style.css -Raw) + "`n--></style>"
  $sig = Get-Content -Path .\html\emailSig.html -Raw
 }
 process {
  $endDate = (Get-Date $_.ad.AccountExpirationDate -f 'dddd, MMMM dd, yyyy')
  $userMsg = $baseMsg -f $_.ad.GivenName, $_.ad.Mail, $_.ad.HomePage, $endDate
  $msgBody = ConvertTo-Html -Head $head -Body $userMsg -PostContent $sig -Title 'CUSD Account Expiring'
  $_.emailParams = @{
   To      = $_.ad.Mail
   Subject = 'CUSD Account Expiring - {0}' -f $endDate
   Body    = $msgBody | Out-String
  }
  $msgData = $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.Mail
  Write-Host ('{0},{1},{2}' -f $msgData ) -F DarkBlue
  Write-Verbose ($MyInvocation.MyCommand.name, $_.emailParams | Out-String )
  Write-Verbose ($MyInvocation.MyCommand.name, $_.emailParams.body | Out-String )
  $_
 }
}

function Format-AdminEmail ([string[]]$manager, [string]$type) {
 begin {
  $mailObj = [PSCustomObject]@{ emailParams = $null }
  $thisDate = (Get-Date -f 'dddd, MMMM dd, yyyy')
  $tableData = @()
  $emailParams = @{
   To      = $manager
   Subject = 'CUSD Accounts {0} - {1}' -f $type, $thisDate
   Body    = $null
  }
 }
 process { $tableData += $_.ad }
 end {
  if (@($tableData).count -lt 1) { return }
  $css = Get-Content -Path '.\html\style.css' -Raw
  $head = '<style TYPE="TEXT/CSS">' + $css + '</style>'
  $message = 'The following employee accounts have been {0}:<br><br>' -f $type
  $sig = Get-Content -Path '.\html\emailSig.html' -Raw
  $html = $tableData | ConvertTo-Html -Property EmployeeID , Name, Mail -Head $head -PreContent $message -PostContent $sig
  $emailParams.Body = $html | Out-String
  Write-Verbose ($emailParams | Out-String )
  Write-Verbose ( $emailParams.Body )
  $mailObj.emailParams = $emailParams
  $mailObj
 }
}

function Format-Object {
 process {
  # Write-Verbose ($_.gettype(), $_.ad.SamAccountName | Out-String)
  [PSCustomObject]@{
   ad          = if ($_.ObjectGUID ) { $_ } else { $null }
   emp         = if ($_.EmpId) { $_ } else { $null }
   emailParams = $null
   expireDate  = $null
  }
 }
}

function Get-ADWithExpirationDate ($days, $ou) {
 $futureDate = (Get-Date).AddDays($days)
 $filter = "
   AccountExpirationDate -like '*' -and
   AccountExpirationDate -le '$futureDate' -and
   Mail -like '*@*' -and
   EmployeeID -like '*'
   "
 $params = @{
  Filter     = $filter
  Properties = 'EmployeeID', 'Mail', 'AccountExpirationDate', 'HomePage', 'info', 'memberof', 'Description'
  SearchBase = $ou
 }
 $results = Get-ADUser @params | Where-Object { $_.info -notmatch 'keep' } | Sort-Object AccountExpirationDate
 Write-Host ( 'AD with expiration date results: {0}' -f @($results).count) -F Green
 $results
}

function Get-ADWithoutExpirationDate {
 process {
  $filter = "EmployeeID -eq '{0}' -and Enabled -eq 'True'" -f $_.emp.EmpID
  $_.ad = Get-ADUser -Filter $filter -Properties AccountExpirationDate, Description, LastLogonDate, Mail | Where-Object {
   ($_.AccountExpirationDate -isnot [datetime] -and $_.accountExpires -isnot [datetime]) -and
   $_.info -notmatch 'keep active'
  }
  if (!$_.ad) { return }
  Write-Verbose ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.emp.EmpID)
  $_
 }
}

function Get-EscapeRetired ($sqlParams, $query) {
 $results = New-SqlOperation @sqlParams -Query $query | ConvertTo-Csv | ConvertFrom-Csv
 Write-Host ( 'Escape Results Count: {0}' -f @($results).count ) -F Green
 $results
}

function New-ExchangeOnlineSession ($userAccount) {
 $sessions = Get-PSSession
 if ($sessions.name -match 'ExchangeOnline') { return }
 Connect-ExchangeOnline -Credential $userAccount -ShowBanner:$false
}

function Remove-ADGroupMemberships {
 process {
  foreach ($group in $_.ad.memberof) {
   $msg = $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.Mail, $group.Split(',')[0].Replace('CN=', '')
   Write-Host ('{0},{1},{2},[{3}]' -f $msg) -F DarkYellow
   Remove-ADGroupMember -Identity $group -Members $_.ad.ObjectGUID -Confirm:$false -WhatIf:$WhatIf
  }
  $_
 }
}

function Remove-EscapeEmailWork ($sqlParams) {
 begin {
  $sql = 'UPDATE HREmployment SET EmailWork = NULL WHERE EmpID = @id'
 }
 process {
  if ($_.ad.EmployeeID -match '[A-Za-z]') {
   Write-Host ('{0},{1},{2},EmployeeID not an integer' -f $MyInvocation.MyCommand.Name, $_.ad.Mail, $_.ad.EmployeeID) -F Red
   return $_
  }
  $sqlVars = ('id={0}' -f $_.ad.EmployeeID)
  Write-Host ('{0},{1},{2},{3},{4}' -f $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.Mail, $sql, $sqlVars) -F Yellow
  if (!$WhatIf) { New-SqlOperation @sqlParams -Query $sql -Parameters $sqlVars }
  $_
 }
}

function Remove-GSuiteAccount {
 begin {
  $gam = '.\bin\gam.exe'
 }
 process {
  $msg = $MyInvocation.MyCommand.Name, $_.ad.EmployeeID, $_.homepage
  $gmail = $_.ad.HomePage
  if ($gmail -notmatch '@') { return $_ }
  Write-Verbose ('{0},{1}' -f $msg[0], "& $gam print users query email:$gmail")
  ($gamUser = & $gam print users query "email:$gmail" | ConvertFrom-Csv) *>$null
  if (!$gamUser) { return (Write-Host ('{0},{1},{2},User Not Found' -f $msg)), $_ }
  Write-Host ('{0},{1}' -f $msg[0], "& $gam delete user $($_.homepage)") -F Blue
  if (!$WhatIf) { (& $gam delete user $gmail) *>$null }
  $_
 }
}

function Remove-StaffADObject {
 process {
  $msgData = $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.Mail, $_.ad.AccountExpirationDate
  Write-Host ('{0},{1},{2},Expired {3}' -f $msgData) -F Yellow
  $removeParams = @{
   Identity    = $_.ad.ObjectGUID
   Confirm     = $False
   ErrorAction = 'SilentlyContinue'
   WhatIf      = $WhatIf
  }
  $objParams = @{
   Filter = "ObjectGUID -eq '{0}'" -f $_.ad.ObjectGUID
  }
  $obj = Get-ADUser @objParams
  Remove-ADUser @removeParams
  if ($WhatIf) { return $_ }
  $obj = Get-ADUser @objParams
  if ($obj) {
   Write-Host ('{0},{1},{2},Delete Failed. Trying -Recursive parameter {3}' -f $msgData) -F Red
   Remove-ADObject @removeParams -Recursive
  }
  $obj = Get-ADUser @objParams
  if ($obj -and !$WhatIf) {
   Write-Host ('{0},{1},{2},Remove-ADUser and Remove-ADObject Failed' -f $msgData) -F Red
   return # Stop here and do not process this account at this time
  }
  $_
 }
}

function Remove-UGNReference ($sqlParams) {
 begin { $query = 'DELETE FROM UGN WHERE SID = @id' }
 process {
  if ($_.ad.EmployeeID -isnot [int]) { return $_ } # Skip this operation if empid is not a number
  $sqlVars = ('id={0}' -f $_.ad.EmployeeID)
  Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.name, $_.ad.Mail, $query, $sqlVars) -F Yellow
  if ($WhatIf) { return $_ }
  New-SqlOperation @sqlParams -Query $query -Parameters $sqlVars
  $_
 }
}

function Remove-FromO365Groups {
 process {
  # Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.Mail) -F Yellow
  $mail = $_.mail
  $exchUser = Get-User -Filter "UserPrincipalName -eq `'$mail`'"
  if (!$exchUser) { return $_ }
  $groups = Get-Recipient -Filter "Members -eq `'$($exchUser.DistinguishedName)`'"
  if (!$groups) { return $_ }
  foreach ($group in $groups) {
   $msg = $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.UserPrincipalName, $group.name
   Write-Host ('{0},{1},{2},{3}' -f $msg) -F Yellow
   $gpParams = @{
    EmailAddresses = @{remove = "$mail" }
    Confirm        = $False
    ErrorAction    = 'SilentlyContinue'
    WhatIf         = $WhatIf
   }
   switch ($group.RecipientTypeDetails) {
    'MailUniversalDistributionGroup' { Get-DistributionGroup $group.name | Set-DistributionGroup @gpParams }
    'MailUniversalSecurityGroup' { Get-DistributionGroup $group.name | Set-DistributionGroup @gpParams }
    'GroupMailbox' { Get-UnifiedGroup $group.name | Set-UnifiedGroup @gpParams }
    default {
     $msgData = $MyInvocation.MyCommand.name, $_.ad.Mail, $group.name, $group.RecipientTypeDetails
     Write-Host ('{0},{1},{2}, Not Sure about group type [{3}]. Skipping.' -f $msgData) -F Red
    }
   }
  }
  $_
 }
}

function Remove-O365Forwarding {
 process {
  $filter = "UserPrincipalName -eq '{0}'" -f $_.ad.Mail
  $mailbox = Get-Mailbox -Filter $filter
  if ($mailbox.ForwardingAddress, $mailbox.ForwardingSmtpAddress -notmatch '@') { return $_ }
  $mailBoxParams = @{
   DeliverToMailboxAndForward = $false
   ForwardingSmtpAddress      = $null
   ForwardingAddress          = $null
   WhatIf                     = $WhatIf
  }
  $msgData = $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.Mail, ($mailbox.ForwardingAddress, $mailbox.ForwardingSmtpAddress -join ',')
  Write-Host ('{0},{1},{2},Removing Mailbox Forwarding [{3}]' -f $msgData ) -F Yellow
  $mailbox | Set-MailBox @mailBoxParams
  $_
 }
}

function Select-Expired {
 begin { $i = 0 }
 process {
  if ($_.ad.AccountExpirationDate -lt [DateTime]::Today -and
   ($_.ad.Enabled -eq 'true' -or $_.ad.Enable -eq $true) ) { $_ }
 }
 end { Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, $i) -f Green }
}

function Select-ExpiredOld ($months) {
 begin { $i = 0 }
 process {
  # Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.ad.Mail, $_.ad.AccountExpirationDate) -F DarkBlue
  if ($_.ad.AccountExpirationDate -le ((Get-Date).AddMonths(-$months)) -and $_.ad.Enabled -eq $False) { $_ }
 }
 end { Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, $i) -f Green }
}

function Select-ExpiringAccounts ([int]$days) {
 begin {
  $i = 0
  function Format-DateToInt { process { [int]$((Get-Date).AddDays($_) | Get-Date -f yyyyMMdd) } }
 }
 process {
  Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.ad.Mail, $_.ad.AccountExpirationDate)
  if (($_.ad.Enable -eq $false) -or ($_.ad.Enable -eq 'false')) { return } # Skip Disabled accounts
  [int]$accExprDate = Get-Date $_.ad.AccountExpirationDate -f yyyyMMdd
  # Compare account expiration date ($accExprDate) to a specific day increment ($checkDate).
  foreach ($checkDate in (1..$days | Format-DateToInt)) {
   # If match then pass input object to next step in pipeline
   if ($checkDate -eq $accExprDate) { $i++; $_ }
  }
 }
 end { Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, $i) -f Green }
}

function Select-StaleSub ($months) {
 begin {
  $staleLogon = (Get-Date).AddMonths(-$months)
 }
 process {
  if ($_.ad.LastLogonDate -gt $staleLogon ) { return }
  Write-Verbose ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.ad.Mail, $_.ad.LastLogonDate)
  $_
 }
}

function Send-ExpireAlert ($emailSender, $bcc) {
 process {
  $params = @{
   To         = $_.emailParams.To
   From       = '<{0}>' -f $emailSender.Username
   Subject    = $_.emailParams.Subject
   Html       = $_.emailParams.Body
   SMTPServer = 'smtp.office365.com'
   Cred       = $emailSender
   UseSSL     = $True
   Port       = 587
   WhatIf     = $WhatIf
   Suppress   = $True
  }
  if ($bcc) { $params += @{Bcc = $bcc } }
  if ($_.emp.EmailHome) { $params.To = $params.To, $_.emp.EmailHome }

  # if (!$WhatIf) { Send-EmailMessage @params }
  Send-EmailMessage @params
  Write-Verbose ($params | Out-String)
  $msg = $MyInvocation.MyCommand.name, ($params.To -join ','), $params.Subject
  Write-Host ('{0},Recipient: [{1}],Subject: [{2}]' -f $msg) -F Green
 }
}

function Set-ADDescription ($msg) {
 process {
  $newDesc = '{0} <{1} {2}>' -f $_.ad.Description, $msg, (Get-Date $_.expireDate -f 'yyyy-MM-dd')
  Write-Host ('{0},{1},[{2}]' -f $MyInvocation.MyCommand.Name, $_.ad.SamAccountName, $newDesc) -f DarkMagenta
  Set-ADUser -Identity $_.ad.ObjectGUID -Description $newDesc -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Set-AeriesSTFInactive ($params) {
 begin { $sql = "UPDATE STF SET tg = 'i' WHERE id = @id" }
 process {
  $sqlVars = "id=$($_.ad.EmployeeID)"
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.Mail) -F DarkYellow
  if ($WhatIf) { return $_ }
  New-SqlOperation @params -Query $sql -Parameters $sqlVars
  $_
 }
}

function Set-RandomPassword {
 process {
  Write-Host('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.Mail) -F DarkYellow
  $securePw = ConvertTo-SecureString -String (New-RandomPassword 20) -AsPlainText -Force
  Set-ADAccountPassword -Identity $_.ad.ObjectGUID -NewPassword $securePw -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Set-ExpireDate ([int]$days) {
 process {
  $lastDay = if ($_.emp.DateTerminationLastDay -match '\d') { $_.emp.DateTerminationLastDay } else { $_.emp.DateTermination }
  $_.expireDate = (Get-Date $lastDay).AddDays($days)
  $msgData = $MyInvocation.MyCommand.name, $_.emp.EmpId, $_.ad.UserPrincipalName, $_.expireDate, $_.ad.LastLogonDate
  Write-Host ('{0},{1},{2},Target Expiration: {3},LastLogon: {4}' -f $msgData) -F Magenta
  Set-ADUser -Identity $_.ad.ObjectGUID -AccountExpirationDate $_.expireDate -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Set-GSuiteArchiveState {
 process {
  Write-Host ('{0},{1},{2} - Archived On' -f $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.HomePage) -F DarkYellow
  if ($WhatIf) { return $_ }
  (& .\bin\gam.exe update user $_.ad.HomePage archived on)*>null
  $_
 }
}

function Show-Object {
 process {
  if ($_) { Write-Verbose ($MyInvocation.MyCommand.name, $_ | Out-String) }
 }
}

function Suspend-GSuiteAccount {
 process {
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.HomePage) -F DarkYellow
  if ($WhatIf) { return $_ }
  (& .\bin\gam.exe update user $_.ad.HomePage suspended on)*>null
  $_
 }
}

function Update-ADStaleSubExpireDate ([int]$days) {
 begin { $expireDate = (Get-Date).AddDays($days) }
 process {
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.ad.SamAccountName, $expireDate) -F Cyan
  Set-ADUser -Identity $_.ad.ObjectGUID -AccountExpirationDate $expireDate -Confirm:$false -WhatIf:$WhatIf
  $_.expireDate = $expireDate
  $_
 }
}

# Main ==========================================================================
Import-Module Mailozaurr -Cmdlet Send-EmailMessage
Import-Module dbatools -Cmdlet Invoke-DbaQuery, Set-DbatoolsConfig, Connect-DbaInstance, Disconnect-DbaInstance
Import-Module ExchangeOnlineManagement -Cmdlet Connect-ExchangeOnline, Disconnect-ExchangeOnline, Get-EXOMailbox
Import-Module CommonScriptFunctions

if ($WhatIf) { Show-TestRun }
Clear-SessionData

$adCmdLets = 'Get-ADUser', 'Set-ADUser', 'Remove-ADObject', 'Remove-ADUser', 'Remove-ADGroupMember', 'Add-ADGroupMember',
'Set-ADAccountPassword', 'Set-ADObject'
Connect-ADSession -DomainControllers $DomainControllers -Cmdlets $adCmdLets -cred $ActiveDirectoryCredential

$empDBParams = @{
 Server     = $EscapeServer
 Database   = $EscapeDatabase
 Credential = $EscapeCredential
}

$SISParams = @{
 Server     = $SISServer
 Database   = $SISDatabase
 Credential = $SISCredential
}

# =========================================================================================
Show-BlockInfo 'Set Stale Sub Expirations'
$escapeRetiredSubs = Get-EscapeRetired $empDBParams (Get-Content -Path '.\sql\retired-subs.sql' -Raw)
$escapeRetiredSubs | Format-Object |
 Get-ADWithoutExpirationDate |
  Select-StaleSub $StaleSubMonths |
   Update-ADStaleSubExpireDate $GracePeriodDays |
    Set-ADDescription 'Expiration Date Set' |
     Show-Object
# =========================================================================================

Show-BlockInfo 'Set Account Expirations'
$escapeRetired = Get-EscapeRetired $empDBParams (Get-Content -Path '.\sql\retired-terminated.sql' -Raw)
$escapeRetired |
 Format-Object |
  Get-ADWithoutExpirationDate |
   Set-ExpireDate $GracePeriodDays |
    Set-ADDescription 'Expiration Date Set' |
     Show-Object

# =========================================================================================
$expireADObjects = Get-ADWithExpirationDate $GracePeriodDays $StaffOU | Format-Object

Show-BlockInfo 'Account Expiration Emails'
$expireADObjects |
 Select-ExpiringAccounts 14 |
  Add-EmpData $empDBParams |
   Format-EmailExpiring |
    Send-ExpireAlert $EmailCredential $BccAddress

Show-BlockInfo 'Disabling Expired Accounts'
New-ExchangeOnlineSession $ExchangeOnlineCredentials
$expireADObjects |
 Select-Expired |
  Disable-ADAccount |
   Set-RandomPassword |
    Clear-ExtensionAttribute1 |
     Remove-ADGroupMemberships |
      Add-DefaultGroups $DefaultGroups |
       Set-AeriesSTFInactive $SISParams |
        Remove-FromO365Groups |
         Remove-O365Forwarding |
          Suspend-GSuiteAccount |
           Set-GSuiteArchiveState |
            Format-AdminEmail $AccountManagerEmail Disabled |
             Send-ExpireAlert $EmailCredential

Disconnect-ExchangeOnline -Confirm:$false

Show-BlockInfo 'Remove Old Expired Accounts'
$expireADObjects |
 Select-ExpiredOld -months $DisabledMaxMonths |
  Remove-UGNReference $sisParams |
   Remove-GSuiteAccount |
    Remove-EscapeEmailWork $empDBParams |
     Disable-PreventAccidentalDelete |
      Remove-StaffADObject |
       Format-AdminEmail $AccountManagerEmail Removed |
              Send-ExpireAlert $EmailCredential

# =========================================================================================

Remove-Module tmp*

if ($WhatIf) { Show-TestRun }