#Requires -Version 5.0
<#
.SYNOPSIS
.DESCRIPTION
 Set Expiration dates on exiting staff accounts, sends emails to accounts with upcoming expiration,
 disables expired accounts, and removes accounts expired for x months.
 An email is sent to users warning them of the impending expiration.
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
 [Parameter(Mandatory = $True)]
 [Alias('DC')]
 [string]$DomainController,
 [Parameter(Mandatory = $True)]
 [Alias('ADCred')]
 [System.Management.Automation.PSCredential]$ActiveDirectoryCredential,
 [Parameter(Mandatory = $True)]
 [string]$StaffOU = 'OU=Employees,OU=Users,OU=Domain_Root,DC=chico,DC=usd',
 [Parameter(Mandatory = $True)]
 [Alias('MSCred')]
 [System.Management.Automation.PSCredential]$MSOLCredentials,
 [Parameter(Mandatory = $True)]
 [Alias('License')]
 [string]$DefaultLicenseSkuID,
 [Parameter(Mandatory = $True)]
 [Alias('AzureCred')]
 [System.Management.Automation.PSCredential]$AzureADCredentials,
 [Parameter(Mandatory = $True)]
 [Alias('ExchCred')]
 [System.Management.Automation.PSCredential]$ExchangeOnlineCredentials,
 [Parameter(Mandatory = $True)]
 [string]$EscapeServer,
 [Parameter(Mandatory = $True)]
 [string]$EscapeDatabase,
 [Parameter(Mandatory = $True)]
 [System.Management.Automation.PSCredential]$EscapeCredential,
 [Parameter(Mandatory = $True)]
 [string]$SISServer,
 [Parameter(Mandatory = $True)]
 [string]$SISDatabase,
 [Parameter(Mandatory = $True)]
 [System.Management.Automation.PSCredential]$SISCredential,
 [Parameter(Mandatory = $True)]
 [Alias('MailCred')]
 [System.Management.Automation.PSCredential]$EmailCredential,
 # Send the heads up email to another account
 [string[]]$BccAddress,
 [Parameter(Mandatory = $True)]
 [string[]]$AccountManagerEmail,
 [Alias('wi')]
 [SWITCH]$WhatIf
)
# ==================================
function Clear-Resources {
 $MyInvocation.MyCommand.name
 Get-Module -name *tmp* | Remove-Module -Confirm:$false -Force -Verbose:$false
 Get-PSSession | Remove-PSSession -Confirm:$false -Verbose:$false
}

function ConvertTo-PSDate {
 process {
  if ($_ -ne [System.DBNull]::Value) {
   Get-Date $_
  }
 }
}

function Disable-ADAccount {
 process {
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.employeeId , $_.mail ) -Fore DarkYellow
  Write-Debug 'Proceed?'
  Set-ADUser -Identity $_.ObjectGUID -Enabled:$false -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Disable-GSuiteForwarding {
 begin {
  $gam = '.\bin\gam-64\gam.exe'
 }
 process {
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.employeeId , $_.homepage ) -Fore DarkYellow
  if ($WhatIf) { Write-Host ("& $gam user {0} forward off" -f $_.homepage) -Fore Blue }
  else { & $gam user $_.homepage forward off }
  $_
 }
}

function Find-MissingExpireDate {
 process {
  # Write-Verbose ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.EmpID )
  $filter = "employeeId -eq `'{0}`'" -f $_.empId
  Get-ADUser -Filter $filter -Properties * | Where-Object {
   $_.AccountExpirationDate -isnot [datetime] -and
   $_.info -notmatch 'keep active' -and
   $_.Enabled -eq $True
  }
 }
}

function Format-DateToInt { process { [int]$((Get-Date).AddDays($_) | Get-Date -f yyyyMMdd) } }

function Format-EmailExpiring {
 begin { $baseMsg = Get-Content -Path '.\lib\EmailExpiring.txt' -Raw }
 process {
  $endDate = (Get-Date $_.AccountExpirationDate -f 'dddd, MMMM dd, yyyy')
  $html = $baseMsg -f $_.givenName, $_.mail, $_.homepage, $endDate
  $emailParams = @{
   to      = $_.mail
   subject = 'CUSD Account Expiring - {0}' -f $endDate
   body    = $html
  }
  Write-Verbose ($emailParams | Out-String )
  Write-Verbose ($emailParams.body | Out-String )
  $emailParams
 }
}

function Format-EmailDisabled {
 begin { $baseMsg = Get-Content -Path '.\lib\EmailDisabled.txt' -Raw }
 process {
  $endDate = (Get-Date -f 'dddd, MMMM dd, yyyy')
  $html = $baseMsg -f $_.name, $_.mail, $_.homepage, $endDate
  $emailParams = @{
   to      = $AccountManagerEmail
   subject = 'CUSD Account Expired - {0} - {1}' -f $_.mail, $endDate
   body    = $html
  }
  Write-Verbose ($emailParams | Out-String )
  Write-Verbose ($emailParams.body | Out-String )
  $emailParams
 }
}

function Format-EmailRemoved {
 begin { $baseMsg = Get-Content -Path '.\lib\EmailRemoved.txt' -Raw }
 process {
  $html = $baseMsg -f $_.name, $_.mail, $_.homepage, (Get-Date -f 'dddd, MMMM dd yyyy')
  $emailParams = @{
   to      = $AccountManagerEmail
   subject = 'CUSD Account Removed - {0}' -f $_.mail
   body    = $html
  }
  Write-Verbose ($emailParams | Out-String )
  Write-Verbose ($emailParams.body | Out-String )
  $emailParams
 }
}

function Get-AccountsWithExpirationDate {
 $msgData = $MyInvocation.MyCommand.Name, $DomainController
 Write-Host ('{0},{1}, Get AD Objects with AccountExpirationDate' -f $msgData) -Fore Green
 $params = @{
  filter     = {
    (AccountExpirationDate -like "*") -and
    (mail -like "*@*") -and
    (employeeId -like "*")
  }
  Properties = 'employeeid', 'mail', 'AccountExpirationDate', 'HomePage', 'info', 'memberof'
  SearchBase = $StaffOU
 }
 $results = Get-ADUser @params | Where-Object {
   (($_.info -notmatch 'keep') -and
   ($_.info -notmatch 'active') -and
   ($_.employeeID -match "\d{4,}"))
 } | Sort-Object Surname

 Write-Host ( 'AD with expiration date results: {0}' -f $results.count) -Fore Green
 $results
}

function Get-EscapeRetired {
 $EscapeParams = @{
  Server     = $EscapeServer
  Database   = $EscapeDatabase
  Credential = $EscapeCredential
  Query      = (Get-Content -Path '.\sql\retired-terminated.sql' -Raw)
 }
 $results = Invoke-Sqlcmd @EscapeParams
 Write-Host ( 'Escape Results Count: {0}' -f $results.count ) -Fore Green
 $results
}

function New-ADSession {
 Write-Host ('{0},{1},Creating AD Domain Controller Session' -f $MyInvocation.MyCommand.Name, $DomainController) -Fore Green
 $adSession = New-PSSession -ComputerName $DomainController -Credential $ActiveDirectoryCredential
 $adCmdLets = 'Get-ADUser', 'Set-ADUser', 'Move-ADObject' , 'Search-ADAccount'
 Import-PSSession -Session $adSession -Module 'ActiveDirectory' -CommandName $adCmdLets -AllowClobber | Out-Null
}

function Remove-ADGroupMemberships {
 process {
  foreach ($group in $_.memberof) {
   Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.name, $_.employeeId, $_.mail, $group) -Fore DarkYellow
   Remove-ADGroupMember -Identity $group -Members $_.ObjectGUID -Confirm:$false -WhatIf:$WhatIf
  }
  $_
 }
}

function Remove-EscapeEmailWork {
 begin {
  $EscapeParams = @{
   Server     = $EscapeServer
   Database   = $EscapeDatabase
   Credential = $EscapeCredential
  }
  $baseSql = 'UPDATE HREmployment SET EmailWork = NULL WHERE EmpID = {0}'
 }
 process {
  $sql = $baseSql -f $_.employeeID
  Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.name, $_.employeeId, $_.mail, $sql) -Fore DarkYellow
  if (-not$WhatIf) { Invoke-Sqlcmd @EscapeParams -Query $sql }
  $_
 }
}

function Remove-GSuiteAccount {
 begin {
  $gam = '.\bin\gam-64\gam.exe'
 }
 process {
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.employeeId, $_.homepage) -Fore DarkYellow
  if ($WhatIf) { Write-Host ("& $gam delete user {0}" -f $_.homepage) }
  else { & $gam delete user $_.homepage }
  $_
 }
}

function Remove-MsolAccount {
 process {
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.employeeId, $_.mail) -Fore DarkYellow
  if (-not$WhatIf) { Get-MSolUser -UserPrincipalName $_.mail | Remove-MsolUser -Force }
  $_
 }
}

function Remove-StaffADObject {
 process {
  $msgData = $MyInvocation.MyCommand.name, $_.employeeId, $_.mail, $_.AccountExpirationDate
  Write-Host ('{0},{1},{2},Expired {3}' -f $msgData) -Fore DarkYellow
  Write-Debug 'Proceed?'
  Remove-ADObject $_.ObjectGUID -Recursive -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Remove-FromO365Groups {
 process {
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.employeeId, $_.mail) -Fore DarkYellow
  $mail = $_.mail
  $msoUser = Get-User -Filter "UserPrincipalName -eq `'$mail`'"
  if ($msoUser) {
   $groups = Get-Recipient -Filter "Members -eq `'$($msoUser.DistinguishedName)`'"
   if ($groups) {
    foreach ($group in $groups) {
     Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.name, $_.employeeId, $_.UserPrincipalName, $group.name) -Fore DarkYellow
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
       $msgData = $MyInvocation.MyCommand.name, $_.mail, $group.name, $group.RecipientTypeDetails
       Write-Host ('{0},{1},{2}, Not Sure about group type [{3}]. Skipping.' -f $msgData) -Fore Red
      }
     }
    }
   }
  }
  $_
 }
}

function Remove-O365Forwarding {
 process {
  $mailBoxParams = @{
   DeliverToMailboxAndForward = $false
   ForwardingSmtpAddress      = $null
   ForwardingAddress          = $null
   WhatIf                     = $WhatIf
  }
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.employeeId, $_.mail) -Fore DarkYellow
  $filter = "UserPrincipalName -eq `'{0}`'" -f $_.mail
  Get-EXOMailbox -Filter $filter | Set-MailBox @mailBoxParams
  $_
 }
}

filter Select-CurrentlyExpired {
 $_ | Where-Object {
  ($null -ne $_.AccountExpirationDate) -and
  ($_.AccountExpirationDate -lt [DateTime]::Today) -and
  ($_.Enabled -eq $True)
 }
}

filter Select-Expired18Months {
 $_ | Where-Object {
  ($null -ne $_.AccountExpirationDate) -and
  ($_.AccountExpirationDate -le ((Get-Date).AddMonths(-18))) -and
  ($_.Enabled -eq $false)
 }
}

function Select-ExpiringAccounts {
 process {
  [int]$accExprDate = Get-Date $_.AccountExpirationDate -f yyyyMMdd
  # Compare account expiration date ($accExprDate) to a specific day increment ($checkDate).
  # foreach ($checkDate in (1, 3, 7, 14 |  Format-DateToInt)) {
  foreach ($checkDate in (1..30 | Format-DateToInt)) {
   # If match then pass input object to next step in pipeline
   if ($checkDate -eq $accExprDate) { $_ }
  }
 }
}

function Send-AlertEmail {
 process {
  $mailParams = @{
   To         = $_.to
   From       = $EmailCredential.Username
   Subject    = $_.subject
   bodyAsHTML = $true
   Body       = $_.body
   SMTPServer = 'smtp.office365.com'
   Cred       = $EmailCredential
   UseSSL     = $True
   Port       = 587
  }
  if ($BccAddress) { $mailParams += @{Bcc = $BccAddress } }

  if (-not$WhatIf) { Send-MailMessage @mailParams }
  $msgData = $MyInvocation.MyCommand.name, ($_.to -join ','), $_.subject
  Write-Host ('{0},Recipient: [{1}],Subject: [{2}]' -f $msgData) -Fore Green
 }
}

function Set-AddressListVisability {
 process {
  $msgData = $MyInvocation.MyCommand.name, $_.employeeId, $_.mail
  Write-Host ('{0},{1},{2},Setting msExchHideFromAddressLists = true' -f $msgData) -Fore DarkYellow
  Set-ADUser -Identity $_.ObjectGUID -Replace @{msExchHideFromAddressLists = $true } -Whatif:$WhatIf
  $_
 }
}

function Set-AeriesSTFStatus {
 begin {
  $SISParams = @{
   Server     = $SISServer
   Database   = $SISDatabase
   Credential = $SISCredential
  }
 }
 process {
  $updateSTF = "UPDATE STF SET tg = `'i`' WHERE id = {0}" -f $_.employeeId
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.mail, $updateSTF)
  if (-not($WhatIf)) { Invoke-SQLCmd @SISParams -Query $updateSTF }
  $_
 }
}

function Set-ExtensionAttribute1 {
 process {
  Write-Host ('{0},{1},{2},Clearing' -f $MyInvocation.MyCommand.name, $_.employeeId, $_.mail)
  Set-ADUser -Identity $_.ObjectGUID -Clear extensionAttribute1 -Whatif:$WhatIf
  $_
 }
}

function Set-MSOLLicense {
 process {
  $upn = $_.UserPrincipalName
  $msolData = Get-MsolUser -UserPrincipalName $upn | Where-Object { $_.IsLicensed -eq $True }
  if ($msolData) {
   if ($msoldata.Licenses.AccountSkuId -ne $DefaultLicenseSkuID) {
    Write-Host ('{0},{1},{2},Removing {3}' -f $upn, $MyInvocation.MyCommand.name, $_.employeeId, ($msoldata.Licenses.AccountSkuId -join ','))
    Write-Host ('{0},{1},{2},Adding {3}' -f $upn, $MyInvocation.MyCommand.name, $_.employeeId, $DefaultLicenseSkuID)
    if (-not$WhatIf) {
     Set-MsolUserLicense -UserPrincipalName $upn -RemoveLicenses $msoldata.Licenses.AccountSkuId
     Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses $DefaultLicenseSkuID
    }
   }
   else {
    $msgData = $MyInvocation.MyCommand.name, $_.employeeId, $upn, $DefaultLicenseSkuID
    Write-Host ('{0},{1},{2},Licensing already correct [{3}]. Skipping.' -f $msgData) -Fore DarkYellow
   }
  }
  else { Write-Host ('{0},{1}, No MSOnline License Present' -f $_.mail, $MyInvocation.MyCommand.name) }
  $_
 }
}

function Set-RandomPassword {
 process {
  Write-Host('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.employeeId, $_.mail) -Fore DarkYellow
  $securePw = ConvertTo-SecureString -String (Get-RandomPassword 20) -AsPlainText -Force
  Set-ADAccountPassword -Identity $_.ObjectGUID -NewPassword $securePw -Confirm:$false -WhatIf:$WhatIf
  $_
 }
}

function Suspend-GSuiteAccount {
 begin {
  $gam = '.\bin\gam-64\gam.exe'
 }
 process {
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.employeeId, $_.homepage) -Fore DarkYellow
  if ($WhatIf) { Write-Host ("& $gam update user {0} suspended on" -f $_.homepage) }
  else { & $gam update user $_.homepage suspended on }
  $_
 }
}

function Set-ExpireDate {
 begin { $expireDate = (Get-Date).AddDays(14) }
 process {
  $msgData = $MyInvocation.MyCommand.name, $_.employeeId, $_.UserPrincipalName, $expireDate
  Write-Host ('{0},{1},{2},{3}' -f $msgData) -Fore DarkYellow
  Set-ADUser -Identity $_.ObjectGUID -AccountExpirationDate $expireDate -Confirm:$false -WhatIf:$WhatIf
 }
}
function Set-AccountExpirations ($escapeData) {
 $escapeData | Find-MissingExpireDate | Set-ExpireDate
}

function Send-ExpirationEmails ($adData) {
 $expiringSoon = $adData | Select-ExpiringAccounts
 Write-Host ( 'Count: {0}' -f $expiringSoon.count ) -Fore Green
 $expiringSoon | Format-EmailExpiring | Send-AlertEmail
}

function Disable-Expired ($adData) {
 Connect-MSOLService -Credential $MSOLCredentials
 Connect-ExchangeOnline -Credential $ExchangeOnlineCredentials
 Connect-AzureAD -Credential $AzureADCredentials
 $expiredAccounts = $adData | Select-CurrentlyExpired
 $verboseList = $expiredAccounts | Select-Object name, AccountExpirationDate, Enabled | Format-Table
 Write-Verbose ($verboseList | Out-String)
 Write-Host ( 'Count: {0}' -f $expiredAccounts.count ) -Fore Green
 $expiredAccounts |
 Disable-ADAccount |
 Set-RandomPassword |
 Set-AddressListVisability |
 Set-ExtensionAttribute1 |
 Remove-ADGroupMemberships |
 Set-AeriesSTFStatus |
 Set-MSOLLicense |
 Remove-FromO365Groups |
 Remove-O365Forwarding |
 Disable-GSuiteForwarding |
 Format-EmailDisabled |
 Send-AlertEmail
}

function Remove-OldAccounts($adData) {
 $expired18Months = $adData | Select-Expired18Months
 $verboseList = $expired18Months | Select-Object Name, AccountExpirationDate, Enabled | Format-Table
 Write-Verbose ($verboseList | Out-String)
 $expired18Months |
 Remove-StaffADObject |
 Remove-MsolAccount |
 Remove-GSuiteAccount |
 Remove-EscapeEmailWork |
 Format-EmailRemoved |
 Send-AlertEmail
}

. .\lib\Clear-SessionData.ps1
. .\lib\Create-O365PSSession.ps1
. .\lib\Get-RandomPassword.ps1
. .\lib\Load-Module.ps1
. .\lib\Show-BlockInfo.ps1
. .\lib\Show-TestRun.ps1

# Main ==========================================================================
Show-TestRun
Clear-SessionData

'MSOnline', 'SQLServer', 'ExchangeOnlineManagement', 'AzureAD' | Load-Module
New-ADSession

$escapeRetired = Get-EscapeRetired
$expiringADAccounts = Get-AccountsWithExpirationDate

Show-BlockInfo 'Set Account Expirations'
Set-AccountExpirations $escapeRetired
Show-BlockInfo 'Send Account Expiration Emails'
Send-ExpirationEmails $expiringADAccounts
Show-BlockInfo 'Disable Expired Accounts'
Disable-Expired $expiringADAccounts
Show-BlockInfo 'Remove Expired Accounts after 18 months'
Remove-OldAccounts $expiringADAccounts

Clear-SessionData
Show-TestRun