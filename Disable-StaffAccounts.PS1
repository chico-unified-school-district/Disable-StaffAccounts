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
  [Parameter(Mandatory = $True)]
  [Alias('DCs')]
  [string[]]$DomainControllers,
  [Parameter(Mandatory = $True)]
  [Alias('ADCred')]
  [System.Management.Automation.PSCredential]$ActiveDirectoryCredential,
  [Parameter(Mandatory = $True)]
  [string]$StaffOU = 'OU=Employees,OU=Users,OU=Domain_Root,DC=chico,DC=usd',
  [Parameter(Mandatory = $True)]
  [Alias('License')]
  [string]$DefaultLicenseSkuID,
  # [Parameter(Mandatory = $True)]
  # [Alias('AzureCred')]
  # [System.Management.Automation.PSCredential]$AzureADCredentials,
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

function New-ExchangeOnlineSession {
  $sessions = Get-PSSession
  if ($sessions.name -match 'ExchangeOnline') { return }
  Connect-ExchangeOnline -Credential $ExchangeOnlineCredentials -ShowBanner:$false
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
    $msgData = $MyInvocation.MyCommand.name, $_.EmployeeID , $_.mail, $_.AccountExpirationDate
    Write-Host ('{0},{1},{2},AccountExpirationDate [{3}]' -f $msgData ) -F DarkYellow
    Write-Debug 'Proceed?'
    Set-ADUser -Identity $_.ObjectGUID -Enabled:$false -Confirm:$false -WhatIf:$WhatIf
    $_
  }
}

function Disable-GSuiteForwarding {
  begin {
    $gam = '.\bin\gam.exe'
  }
  process {
    # TODO scope - access denied issue - gamadv might work instead
    Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.EmployeeID , $_.homepage ) -F DarkYellow
    Write-Host ("& $gam user {0} forward off" -f $_.homepage) -F Yello
    if (!$WhatIf) { & $gam user $_.homepage forward off }
    $_
  }
}

function Find-MissingExpireDate {
  process {
    # Write-Verbose ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.EmpID )
    $filter = "employeeId -eq `'{0}`'" -f $_.empId
    Get-ADUser -Filter $filter -Properties * | Where-Object {
      ($_.AccountExpirationDate -isnot [datetime] -and $_.accountExpires -isnot [datetime]) -and
      $_.info -notmatch 'keep active' -and
      $_.Enabled -eq $True
    }
  }
}

function Format-DateToInt { process { [int]$((Get-Date).AddDays($_) | Get-Date -f yyyyMMdd) } }

function Format-EmailExpiring {
  begin { $baseMsg = Get-Content -Path .\html\EmailExpiring.html -Raw }
  process {
    $endDate = (Get-Date $_.AccountExpirationDate -f 'dddd, MMMM dd, yyyy')
    $html = $baseMsg -f $_.givenName, $_.mail, $_.homepage, $endDate
    $emailParams = @{
      to      = $_.mail
      subject = 'CUSD Account Expiring - {0}' -f $endDate
      body    = $html
    }
    $msgData = $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail
    Write-Host ('{0},{1},{2}' -f $msgData ) -F DarkBlue
    Write-Verbose ($emailParams | Out-String )
    # Write-Verbose ($emailParams.body | Out-String )
    $emailParams
  }
}

function Format-Email ($type) {
  begin {
    $thisDate = (Get-Date -f 'dddd, MMMM dd, yyyy')
    $tableData = @()
    $emailParams = @{
      to      = $AccountManagerEmail
      subject = 'CUSD Accounts {0} - {1}' -f $type, $thisDate
    }
  }
  process { $tableData += $_ }
  end {
    if ($tableData.count -lt 1) { return }
    $css = Get-Content -Path '.\html\style.css' -Raw
    $head = '<style TYPE="TEXT/CSS">' + $css + '</style>'
    $message = 'The following employee accounts have been {0}:<br><br>' -F $type
    $sig = Get-Content -Path '.\html\emailSig.html' -Raw
    $html = $tableData | ConvertTo-Html -Property EmployeeId , Name, Mail -Head $head -PreContent $message -PostContent $sig
    $emailParams.body = $html | Out-String
    Write-Verbose ( $emailParams.Body )
    Write-Verbose ($emailParams | Out-String )
    $emailParams
  }
}

function Get-AccountsWithExpirationDate {
  Write-Host ('{0}' -f $MyInvocation.MyCommand.Name) -F Green
  $params = @{
    filter     = {
    (AccountExpirationDate -like "*") -and
    (mail -like "*@*") -and
    (employeeId -like "*")
    }
    Properties = 'employeeid', 'mail', 'AccountExpirationDate', 'homepage', 'info', 'memberof'
    SearchBase = $StaffOU
  }
  $results = Get-ADUser @params | Where-Object {
   (($_.info -notmatch 'keep') -and
   ($_.info -notmatch 'active') -and
   ($_.EmployeeID -match "\d{4,}"))
  } | Sort-Object Surname

  Write-Host ( 'AD with expiration date results: {0}' -f @($results).count) -F Green
  $results
}

function Get-EscapeRetired ($sqlParams, $query) {
  $results = New-SqlOperation @sqlParams -Query $query
  Write-Host ( 'Escape Results Count: {0}' -f @($results).count ) -F Green
  $results
}

function New-ADSessionOld {
  Write-Host ('{0},{1},Creating AD Domain Controller Session' -f $MyInvocation.MyCommand.Name, $DomainControllers) -F Green
  $adSession = New-PSSession -ComputerName $DomainControllers -Credential $ActiveDirectoryCredential
  $adCmdLets = 'Get-ADUser', 'Set-ADUser', 'Remove-ADObject' , 'Remove-ADGroupMember', 'Set-ADAccountPassword'
  Import-PSSession -Session $adSession -Module 'ActiveDirectory' -CommandName $adCmdLets -AllowClobber | Out-Null
}

function Remove-ADGroupMemberships {
  process {
    foreach ($group in $_.memberof) {
      Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail, $group) -F DarkYellow
      Remove-ADGroupMember -Identity $group -Members $_.ObjectGUID -Confirm:$false -WhatIf:$WhatIf
    }
    $_
  }
}

function Remove-EscapeEmailWork ($sqlParams) {
  begin { $query = 'UPDATE HREmployment SET EmailWork = NULL WHERE EmpID = @id' }
  process {
    if (!$_.EmployeeID) { return (Write-Error '{0},{1},Missing critical data' -f $MyInvocation.MyCommand.Name) }
    $sqlVars = ('id={0}' -f $_.EmployeeID)
    Write-Host ('{0},{1},{2},{3},{4}' -f $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail, $query, $sqlVars) -F Yellow
    if (!$WhatIf) { New-SqlOperation @sqlParams -Query $query -Parameters $sqlVars }
    $_
  }
}

function Remove-GSuiteAccount {
  begin {
    $gam = '.\bin\gam.exe'
  }
  process {
    $msg = $MyInvocation.MyCommand.Name, $_.EmployeeID, $_.homepage
    $gmail = $_.homepage
    Write-Verbose ('{0},{1}' -f $msg[0], "& $gam print users query email:$gmail")
    ($gamUser = & $gam print users query "email:$gmail" | ConvertFrom-Csv) *>$null
    if (!$gamUser) { return (Write-Host ("{0},{1},{2},User Not Found" -f $msg)), $_ }
    Write-Host ("{0},{1}" -f $msg[0], "& $gam delete user $($_.homepage)") -F Blue
    if (!$WhatIf) { (& $gam delete user $_.homepage) *>$null }
    $_
  }
}

function Remove-StaffADObject {
  process {
    $msgData = $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail, $_.AccountExpirationDate
    Write-Host ('{0},{1},{2},Expired {3}' -f $msgData) -F Yellow
    Write-Debug 'Proceed?'
    $removeParams = @{
      Identity    = $_.ObjectGUID
      Confirm     = $False
      ErrorAction = 'SilentlyContinue'
      WhatIf      = $WhatIf
    }
    $objParams = @{
      Filter = "ObjectGUID -eq '{0}'" -f $_.ObjectGUID
    }
    $obj = Get-ADUser @objParams
    Remove-ADUser @removeParams
    $obj = Get-ADUser @objParams
    if ($obj) {
      Write-Host ('{0},{1},{2},Delete Failed. Trying -Recursive parameter {3}' -f $msgData) -F Red
      Remove-ADObject @removeParams -Recursive
    }
    $obj = Get-ADUser @objParams
    if ($obj -and !$WhatIf) {
      Write-Host ('{0},{1},{2},Remove-ADUser and Remove-ADObject Failed' -f $msgData) -F Red
      return
    }
    $_
  }
}

function Remove-UGNReference ($sqlParams) {
  begin { $query = 'DELETE FROM UGN WHERE SID = @id' }
  process {
    $sqlVars = ('id={0}' -f $_.EmployeeID)
    Write-Host ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.name, $_.mail, $query, $sqlVars) -F Yellow
    if (!($WhatIf)) { New-SqlOperation @sqlParams -Query $removeFromUGN -Parameters $sqlVars }
    $_
  }
}

function Remove-FromO365Groups {
  process {
    Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail) -F Yellow
    $mail = $_.mail
    $exchUser = Get-User -Filter "UserPrincipalName -eq `'$mail`'"
    if ($exchUser) {
      $groups = Get-Recipient -Filter "Members -eq `'$($exchUser.DistinguishedName)`'"
      if ($groups) {
        foreach ($group in $groups) {
          $msg = $MyInvocation.MyCommand.name, $_.EmployeeID, $_.UserPrincipalName, $group.name
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
              $msgData = $MyInvocation.MyCommand.name, $_.mail, $group.name, $group.RecipientTypeDetails
              Write-Host ('{0},{1},{2}, Not Sure about group type [{3}]. Skipping.' -f $msgData) -F Red
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
    Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail) -F Yellow
    $mailBoxParams = @{
      DeliverToMailboxAndForward = $false
      # WarningAction              = 'Ignore'
      # WarningVariable            = 'setExoWarn'
      WhatIf                     = $WhatIf
    }
    $filter = "UserPrincipalName -eq `'{0}`'" -f $_.mail
    $mailbox = Get-Mailbox -Filter $filter
    if ($mailbox.ForwardingAddress -match '[A-Za-z0-9]') {
      $msgData = $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail, $mailbox.ForwardingAddress
      Write-Host ('{0},{1},{2},Removing [{3}]' -f $msgData ) -F Yellow
      $mailBoxParams.ForwardingAddress = $null
      Get-EXOMailbox -Filter $filter | Set-MailBox @mailBoxParams
    }
    if ($mailbox.ForwardingSmtpAddress -match '[A-Za-z0-9]') {
      $msgData = $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail, $mailbox.ForwardingSmtpAddress
      Write-Host ('{0},{1},{2},Removing [{3}]' -f $msgData ) -F Yellow
      $mailBoxParams.ForwardingSmtpAddress = $null
      Get-EXOMailbox -Filter $filter | Set-MailBox @mailBoxParams
    }
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
      From       = '<{0}>' -f $EmailCredential.Username
      Subject    = $_.subject
      BodyAsHTML = $true
      Body       = $_.body | Out-String
      SMTPServer = 'smtp.office365.com'
      Cred       = $EmailCredential
      UseSSL     = $True
      Port       = 587
    }
    if ($BccAddress) { $mailParams += @{Bcc = $BccAddress } }

    if (!$WhatIf) { Send-MailMessage @mailParams }
    Write-Verbose ($mailParams | Out-String)
    $logData = $MyInvocation.MyCommand.name, ($_.to -join ','), $_.subject
    Write-Host ('{0},Recipient: [{1}],Subject: [{2}]' -f $logData) -F Green
  }
}

function Set-AddressListVisibility {
  process {
    $msgData = $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail
    Write-Host ('{0},{1},{2},Setting msExchHideFromAddressLists = true' -f $msgData) -F DarkYellow
    Set-ADUser -Identity $_.ObjectGUID -Replace @{msExchHideFromAddressLists = $true } -Whatif:$WhatIf
    $_
  }
}

function Set-AeriesSTFStatus ($sqlParams) {
  begin { $query = "UPDATE STF SET tg = `'i`' WHERE id = @id" }
  process {
    $sqlVars = ('id={0}' -f $_.EmployeeID)
    Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.mail, $query, $sqlVars)
    if (!($WhatIf)) { New-SqlOperation @sqlParams -Query $query -Parameters $sqlVars }
    $_
  }
}

function Set-ExtensionAttribute1 {
  process {
    Write-Host ('{0},{1},{2},Clearing' -f $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail)
    Set-ADUser -Identity $_.ObjectGUID -Clear extensionAttribute1 -Whatif:$WhatIf
    $_
  }
}

function Set-RandomPassword {
  process {
    Write-Host('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.EmployeeID, $_.mail) -F DarkYellow
    $securePw = ConvertTo-SecureString -String (New-RandomPassword 20) -AsPlainText -Force
    Set-ADAccountPassword -Identity $_.ObjectGUID -NewPassword $securePw -Confirm:$false -WhatIf:$WhatIf
    $_
  }
}

function Suspend-GSuiteAccount {
  begin {
    $gam = '.\bin\gam.exe'
  }
  process {
    Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.EmployeeID, $_.homepage) -F DarkYellow
    Write-Host ("{0},& $gam update user {0} suspended on" -f $MyInvocation.MyCommand.name, $_.homepage) -F Yellow
    if (!$WhatIf) {
   ($result = & $gam update user $_.homepage suspended on)*>null
      Write-Host $result -F DarkCyan
    }
    $_
  }
}

function Set-ExpireDate {
  process {
    $id = $_.EmployeeID
    $escUser = $global:escapeRetired.Where({ [int]$_.empId -eq [int]$id })
    $expireDate = (Get-Date).AddDays(30)
    $dtld = $escUser.DateTerminationLastDay
    $dt = $escUser.DateTermination
    if ($dtld -match '\d') { $expireDate = Get-Date $dtld }
    if ($dt -match '\d') {
      $date2 = Get-Date $dt
      if ($date2 -gt $date1) {
        $expireDate = $date2
      }
    }
    $expireDate2 = (Get-Date $expireDate).AddDays(30)
    $msgData = $MyInvocation.MyCommand.name, $_.EmployeeID, $_.UserPrincipalName, $expireDate2, $_.LastLogonDate
    Write-Host ('{0},{1},{2},Target Expiration: {3},LastLogon: {4}' -f $msgData) -F DarkYellow
    $setUser = @{
      Identity = $_.ObjectGUID
      Confirm  = $False
      WhatIf   = $WhatIf
    }
    Set-ADUser @setUser -AccountExpirationDate $expireDate2
  }
}

function Set-AccountExpirations {
  $global:escapeRetired | Find-MissingExpireDate | Set-ExpireDate
}

function Send-ExpirationEmails ($adData) {
  $expiringSoon = $adData | Select-ExpiringAccounts
  $i = ($expiringSoon | Measure-Object).count
  Write-Host ( '{0},Count: {1}' -f $MyInvocation.MyCommand.Name, $i ) -F Green
  $expiringSoon | Format-EmailExpiring | Send-AlertEmail
}

function Disable-Expired ($adData, $SISParams) {
  New-ExchangeOnlineSession
  # Connect-AzureAD -Credential $AzureADCredentials
  $expiredAccounts = $adData | Select-CurrentlyExpired
  $verboseList = $expiredAccounts | Select-Object name, AccountExpirationDate, Enabled | Format-Table
  Write-Verbose ($verboseList | Out-String)
  Write-Host ( '{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($expiredAccounts).count ) -F Green

  $expiredAccounts |
  # $users = @(); $users += ad Testrock -All; $users += ad Testerman -All ; $users |
  Disable-ADAccount |
  Set-RandomPassword |
  Set-AddressListVisibility |
  Set-ExtensionAttribute1 |
  Remove-ADGroupMemberships |
  Set-AeriesSTFStatus $SISParams |
  Remove-FromO365Groups |
  Remove-O365Forwarding |
  Suspend-GSuiteAccount |
  Format-Email 'Disabled' |
  Send-AlertEmail

  Disconnect-ExchangeOnline -Confirm:$false
}

function Disable-PreventAccidentalDelete {
  begin {
  }
  process {
    Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.mail) -F Yellow
    $prevParams = @{
      Identity                        = $_.ObjectGUID
      ProtectedFromAccidentalDeletion = $False
      Confirm                         = $False
      ErrorAction                     = 'Stop'
      WhatIf                          = $WhatIf
    }
    Set-ADObject @prevParams
    $_
  }
}

function Remove-OldAccounts($adData, $empParams, $sisParams) {
  $expired18Months = $adData | Select-Expired18Months
  $verboseList = $expired18Months | Select-Object Name, AccountExpirationDate, Enabled | Format-Table
  Write-Verbose ( $verboseList | Out-String )
  Write-Host ( '{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($expired18Months).count ) -F Green

  $expired18Months |
  # $users = @(); $users += ad Testrock -All; $users += ad Testerman -All ; $users |
  Remove-UGNReference $sisParams |
  Remove-GSuiteAccount |
  Remove-EscapeEmailWork $empParams |
  Disable-PreventAccidentalDelete |
  Remove-StaffADObject |
  Format-Email Removed |
  Send-AlertEmail
}

Import-Module -Name dbatools, ExchangeOnlineManagement, CommonScriptFunctions

# Imported Functions
. .\lib\Create-O365PSSession.ps1

# Main ==========================================================================

if (!$WhatIf) { Show-TestRun }
Clear-SessionData

$dc = Select-DomainController $DomainControllers
$adCmdLets = 'Get-ADUser',
'Set-ADUser',
'Remove-ADObject',
'Remove-ADUser',
'Remove-ADGroupMember',
'Set-ADAccountPassword',
'Set-ADObject'
New-ADSession -dc $dc -cmdlets $adCmdLets -cred $ActiveDirectoryCredential

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

$global:escapeRetired = Get-EscapeRetired $empDBParams (Get-Content -Path '.\sql\retired-terminated.sql' -Raw)
$expiringADAccounts = Get-AccountsWithExpirationDate

Show-BlockInfo 'Setting Account Expirations'
Set-AccountExpirations
Show-BlockInfo 'Sending Account Expiration Emails'
Send-ExpirationEmails $expiringADAccounts
Show-BlockInfo 'Disabling Expired Accounts'
Disable-Expired $expiringADAccounts $SISParams
Show-BlockInfo 'Removing Expired Accounts after 18 months'
Remove-OldAccounts $expiringADAccounts $empDBParams $SISParams

Clear-SessionData
if (!$WhatIf) { Show-TestRun }