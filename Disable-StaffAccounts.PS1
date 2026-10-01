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
 [Parameter(Mandatory = $True)][Alias('ADCred')][PSCredential]$ADCredential,
 [Parameter(Mandatory = $false)][string[]]$BccAddress,
 [Parameter(Mandatory = $True)][Alias('License')][string[]]$DefaultGroups,
 [Parameter(Mandatory = $True)][Alias('MailCred')][PSCredential]$EmailCredential,
 [Parameter(Mandatory = $True)][PSCredential]$EmpCredential,
 [Parameter(Mandatory = $True)][string]$EmpDatabase,
 [Parameter(Mandatory = $True)][string]$EmpServer,
 [Parameter(Mandatory = $True)][Alias('ExchCred')][PSCredential]$ExchangeOnlineCredentials,
 [Parameter(Mandatory = $True)][int]$GracePeriodDays,
 [Parameter(Mandatory = $True)][int]$DisabledMaxMonths,
 [Parameter(Mandatory = $True)][string]$OrgUnitNoLicAD,
 [Parameter(Mandatory = $True)][string]$OrgUnitNoLicGoogle,
 [Parameter(Mandatory = $True)][string[]]$ProtectedGroups,
 [Parameter(Mandatory = $True)][PSCredential]$SISCredential,
 [Parameter(Mandatory = $True)][string]$SISDatabase,
 [Parameter(Mandatory = $True)][string]$SISServer,
 # [Parameter(Mandatory = $True)][string]$StaffOU,
 [Parameter(Mandatory = $True)][int]$StaleSubMonths,
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

function Clear-ExtensionAttributes ([pscredential]$cred) {
 process {
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.info) -F DarkYellow
  # 1. Generate the array of all 15 attribute names
  $AllExtensions = 1..15 | ForEach-Object { "extensionAttribute$_" }
  # 2. Clear them all at once for the target user
  Set-ADUser -Identity $_.ad.ObjectGUID -Clear $AllExtensions -Credential $cred -WhatIf:$WhatIf
  $_
 }
}

function Convert-ToSharedMailbox {
 process {
  if (!($_.exchangeData)) { return $_ } # bypass if no exchange data (not licensed for O365)
  $msgData = $MyInvocation.MyCommand.Name, $_.info
  if ($_.exchangeData.IsShared -eq $true) { return $_ } # Already Set to Shared
  Set-Mailbox -Identity $_.ad.UserPrincipalName -Type Shared -WhatIf:$WhatIf # Convert to Shared Mailbox
  if (!$WhatIf) {
   Write-Host ('{0},{1},Waiting for mailbox type change to propagate...' -f $msgData) -F Blue
   Start-Sleep -Seconds 30
  }
  $params = @{
   Filter = "UserPrincipalName -eq '{0}'" -f $_.ad.UserPrincipalName
  }
  if ((Get-Mailbox @params).IsShared -ne $true) {
   Write-Warning ('{0},{1},NOT Converted to Shared Mailbox.' -f $msgData)
   if (!$WhatIf) { return } # Skip further processing
  }
  Write-Host ('{0},{1},Converted to Shared Mailbox' -f $msgData) -F Green
  $_
 }
}

function Disable-ADAccount ([pscredential]$cred) {
 process {
  $msgData = $MyInvocation.MyCommand.name, $_.info
  Write-Host ('{0},{1}' -f $msgData ) -F DarkYellow
  $params = @{
   Identity   = $_.ad.ObjectGUID
   Enabled    = $false
   Credential = $cred
   Confirm    = $false
   WhatIf     = $WhatIf
  }
  Set-ADUser @params
  $_
 }
}

function Disable-PreventAccidentalDelete ([pscredential]$cred) {
 process {
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.Name, $_.info) -F Yellow
  $prevParams = @{
   Identity                        = $_.ad.ObjectGUID
   ProtectedFromAccidentalDeletion = $False
   Confirm                         = $False
   Credential                      = $cred
   ErrorAction                     = 'Stop'
   WhatIf                          = $WhatIf
  }
  Set-ADObject @prevParams
  $_
 }
}

function Disable-GoogleForwarding {
 process {
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.info ) -F Blue
  # Write-Host ("& $gam user {0} forward off" -f $_.ad.homepage) -F Yellow
  if (!$WhatIf) { (& $gam user "$($_.ad.homepage)" forward off)*>$null }
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
  # if (!($_.exchangeData)) { return } # Skip accounts without exchange data (not licensed for O365)
  $endDate = (Get-Date $_.ad.AccountExpirationDate -f 'dddd, MMMM dd, yyyy')
  $userMsg = $baseMsg -f $_.ad.GivenName, $_.ad.Mail, $_.ad.HomePage, $endDate
  $msgBody = ConvertTo-Html -Head $head -Body $userMsg -PostContent $sig -Title 'CUSD Account Expiring'
  $_.emailParams = @{
   To      = $_.ad.Mail
   Subject = 'CUSD Account Expiring - {0}' -f $endDate
   Body    = $msgBody | Out-String
  }
  $msgData = $MyInvocation.MyCommand.name, $_.info
  Write-Host ('{0},{1}' -f $msgData ) -F DarkBlue
  Write-Verbose ($MyInvocation.MyCommand.name, $_.emailParams | Out-String )
  Write-Verbose ($MyInvocation.MyCommand.name, $_.emailParams.body | Out-String )
  $_
 }
}

function Format-Object {
 process {
  [PSCustomObject]@{
   ad           = if ($_.ObjectGUID ) { $_ } else { $null }
   exchangeData = $null
   emp          = if ($_.EmpId) { $_ } else { $null }
   emailParams  = $null
   expireDate   = $null
   GoogleData   = $null
   info         = $null
  }
 }
}

function Get-ADWithExpirationDate ($days, $props, [pscredential]$cred) {
 $futureDate = (Get-Date).AddDays($days)
 # "(-not(employeeType -eq 'student'))" works, while "EmployeeType -ne 'student'" can fail due to null value
 $filter = "
   AccountExpirationDate -like '*' -and
   AccountExpirationDate -le '$futureDate' -and
   Mail -like '*@*' -and
   EmployeeID -like '*' -and
   -not(info -like '*keep*') -and
   (-not(employeeType -eq 'student'))
   "
 $params = @{
  Filter     = $filter
  Properties = $props
  Credential = $cred
 }
 $results = Get-ADUser @params
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($results).count) -F Green
 $results
}

function Get-ADWithoutExpirationDate ($props, [pscredential]$cred) {
 $filter = "
 accountExpires -eq '9223372036854775807' -and
 Enabled -eq 'True' -and
 mail -like '*' -and
 EmployeeID -like '*' -and
 -not(info -like '*keep*') -and
 (-not(employeeType -eq 'student'))
  "
 $params = @{
  Filter     = $filter
  Properties = $props
  Credential = $cred
 }
 $results = Get-ADUser @params
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($results).count) -F Green
 $results
}

function Get-ADData ($props, [pscredential]$cred) {
 $filter = "
 mail -like '*' -and
 EmployeeID -like '*' -and
 -not(info -like '*keep*') -and
 (-not(employeeType -eq 'student'))
  "
 $params = @{
  Filter     = $filter
  Properties = $props
  Credential = $cred
 }
 $results = Get-ADUser @params
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($results).count) -F Green
 $results | Sort-Object -Property LastLogonDate
}

function Get-EmpData ($sqlParams) {
 $sql = "SELECT EmpId, EmailHome from vwHREmploymentList WHERE EmailWork LIKE '%@%'"
 $results = New-SqlOperation @sqlParams -Query $sql | ConvertTo-Csv | ConvertFrom-Csv
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($results).count) -F Green
 $results
}

function Get-EmpRetired ($sqlParams, $query) {
 $results = New-SqlOperation @sqlParams -Query $query | ConvertTo-Csv | ConvertFrom-Csv
 Write-Host ( '{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($results).count ) -F Green
 $results
}

function Get-ExchangeData {
 Write-Host ('{0},May take some time...' -f $MyInvocation.MyCommand.Name) -f Yellow
 $params = @{
  ResultSize = 'Unlimited'
  Properties = 'IsShared', 'ForwardingAddress', 'ForwardingSmtpAddress'
  Verbose    = $false
 }
 $results = Get-EXOMailbox @params
 Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, @($results).count) -F Green
 $results
}

function New-ExchangeOnlineSession ($userAccount) {
 $sessions = Get-PSSession
 if ($sessions.name -match 'ExchangeOnline') { return }
 Connect-ExchangeOnline -Credential $userAccount -ShowBanner:$false
}

function Remove-ADGroupMemberships ($exGroups, [pscredential]$cred) {
 process {
  $groups = @(($_.ad.memberOf)).ForEach({ $_.Split(',')[0].Replace('CN=', '') })
  foreach ($group in $groups) {
   if ($exGroups -match $group) { continue }
   $params = @{
    Identity   = $group
    Members    = $_.ad.ObjectGUID
    Confirm    = $false
    Credential = $cred
    WhatIf     = $WhatIf
   }
   Remove-ADGroupMember @params
   $msg = $MyInvocation.MyCommand.name, $_.info, $group
   Write-Host ('{0},{1},[{2}]' -f $msg) -F Blue
  }
  $_
 }
}

function Remove-EmpEmailWork ($sqlParams) {
 begin {
  $sql = 'UPDATE HREmployment SET EmailWork = NULL WHERE EmpID = @id'
 }
 process {
  if ($_.ad.EmployeeID -match '[A-Za-z]') {
   Write-Host ('{0},{1}, EmployeeID not an integer [{2}]' -f $MyInvocation.MyCommand.Name, $_.info, $_.ad.EmployeeID) -F Red
   return $_
  }
  $sqlVars = ('id={0}' -f $_.ad.EmployeeID)
  # Write-Verbose ('{0},{1},{2},{3},{4}' -f $MyInvocation.MyCommand.name, $_.ad.EmployeeID, $_.ad.Mail, $sql, $sqlVars)
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.info) -F Yellow
  if (!$WhatIf) { New-SqlOperation @sqlParams -Query $sql -Parameters $sqlVars }
  $_
 }
}

function Remove-GoogleAccount {
 process {
  if ($gmail -notmatch '@') { return $_ }
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.Name, $_.info, $_.ad.HomePage) -F Blue
  if (!$WhatIf) {
   try { (& $gam delete user "$($_.ad.HomePage)") *>$null }
   catch {
    Write-Host ('{0},{1},Error Deleting: {2}' -f $MyInvocation.MyCommand.Name, $_.info, $_.ad.HomePage) -F Red
   }
  }
  $_
 }
}

function Remove-GoogleLicense ($ou) {
 begin {
  $license = @(
   # 1010310009 # Google Workspace for Education Plus (Staff)
   1010310008 # Google Workspace for Education Plus
  )
 }
 process {
  if (!($_.googleData)) { return $_ } # Skip if no Google data
  $i = 20
  do {
   # Wait for Google Workspace to update user orgUnit
   ($ouCheck = & $gam print users query "email:$($_.ad.HomePage)" fields 'orgUnitPath' | ConvertFrom-Csv)*>$null
   if (!$WhatIf -and !$ouCheck) { Start-Sleep 7 }
   $i--
  } until ($WhatIf -or $ouCheck.orgUnitPath -match $ou -or ($i -eq 0))

  $ErrorActionPreference = 'SilentlyContinue'

  foreach ($lic in $license) {
   $msgData = $MyInvocation.MyCommand.name, $_.info, $_.ad.HomePage, $lic
   Write-Host ('{0},{1},{2},Removing License [{3}]' -f $msgData) -F DarkMagenta
   if (!$WhatIf) {
    try { (& $gam user "$($_.ad.HomePage)" del license $lic)*>null }
    catch {
     Write-Host ('{0},{1},Error Removing License [{2}] : {3}' -f $MyInvocation.MyCommand.Name, $_.info, $lic) -F Red
    }
   }
  }

  $ErrorActionPreference = 'Stop'
  $_
 }
}

function Remove-StaffADObject ([pscredential]$cred) {
 process {
  $msgData = $MyInvocation.MyCommand.name, $_.info
  Write-Host ('{0},{1}' -f $msgData) -f Yellow
  $removeParams = @{
   Identity    = $_.ad.ObjectGUID
   Confirm     = $False
   Credential  = $cred
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
   Write-Host ('{0},{1},Delete Failed. Trying -Recursive parameter.' -f $msgData) -F Red
   Remove-ADObject @removeParams -Recursive
  }
  $obj = Get-ADUser @objParams
  if ($obj -and !$WhatIf) {
   Write-Host ('{0},{1},Remove-ADUser and Remove-ADObject Failed' -f $msgData) -F Red
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
  Write-Verbose ('{0},{1},{2},{3}' -f $MyInvocation.MyCommand.name, $_.ad.Mail, $query, $sqlVars)
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.info) -F Yellow
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
   $msg = $MyInvocation.MyCommand.name, $_.info, $group.name
   Write-Host ('{0},{1},{2}' -f $msg) -F Yellow
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
     $msgData = $MyInvocation.MyCommand.name, $_.info, $group.name, $group.RecipientTypeDetails
     Write-Host ('{0},{1},{2}, Not Sure about group type [{3}]. Skipping.' -f $msgData) -F Red
    }
   }
  }
  $_
 }
}

function Remove-O365Forwarding {
 process {
  if (!($_.exchangeData)) { return $_ } # bypass if no exchange data (not licensed for O365)
  $filter = "UserPrincipalName -eq '{0}'" -f $_.ad.Mail
  $mailbox = Get-EXOMailbox -Filter $filter
  if ($mailbox.ForwardingAddress, $mailbox.ForwardingSmtpAddress -notmatch '@') { return $_ }
  $mailBoxParams = @{
   DeliverToMailboxAndForward = $false
   ForwardingSmtpAddress      = $null
   ForwardingAddress          = $null
   WhatIf                     = $WhatIf
  }
  $msgData = $MyInvocation.MyCommand.name, $_.info, ($mailbox.ForwardingAddress, $mailbox.ForwardingSmtpAddress -join ',')
  Write-Host ('{0},{1},Removing Mailbox Forwarding [{2}]' -f $msgData ) -F Blue
  $mailbox | Set-MailBox @mailBoxParams
  $_
 }
}

function Remove-SharedMailbox {
 process {
  if (!($_.exchangeData)) { return $_ } # bypass if no exchange data (not licensed for O365)
  $msgData = $MyInvocation.MyCommand.Name, $_.info
  $params = @{
   Filter = "UserPrincipalName -eq '{0}'" -f $_.ad.UserPrincipalName
  }
  if ((Get-EXOMailbox @params).IsShared -ne $true) { return $_ } # Not a shared mailbox
  if (!(Get-EXOMailbox @params)) {
   Write-Warning ('{0},{1},Mailbox Not Found.' -f $msgData)
   return $_
  }
  Remove-Mailbox -Identity $_.ad.UserPrincipalName -Confirm:$false -WhatIf:$WhatIf
  if ($WhatIf) { return $_ }
  Start-Sleep -Seconds 10 # Allow time for mailbox deletion to propagate
  if (Get-ExMailbox @params) {
   Write-Warning ('{0},{1},NOT Deleted.' -f $msgData)
   return # Skip further processing
  }
  Write-Host ('{0},{1},Shared Mailbox Deleted' -f $msgData) -F Green
  $_
 }
}

function Select-ADExpired {
 process {
  $_ | Where-Object { $_.ad.AccountExpirationDate -le ([DateTime]::Today) }
 }
}

function Select-ADExpiredOld ($months) {
 begin { $i = 0 }
 process {
  if ($_.ad.AccountExpirationDate -le ((Get-Date).AddMonths(-$months)) -and $_.ad.Enabled -eq $False) {
   $i++
   $_
  }
 }
 end { Write-Host ('{0},Count: {1}' -f $MyInvocation.MyCommand.Name, $i) -f Green }
}

function Select-ADExpiringSoon ($days) {
 begin {
  $today = [datetime]::Today
  $cutOffData = $today.AddDays($days)
 }
 process {
  $expDate = $_.ad.AccountExpirationDate
  $_ | Where-Object { $expDate -gt $today -and $expDate -lt $cutOffData }
 }
}

function Select-ADWithExpiration {
 process {
  $_ | Where-Object { $_.ad.AccountExpirationDate -is [datetime] }
 }
}

function Select-ADWithoutExpiration {
 process {
  $_ | Where-Object { $_.ad.AccountExpirationDate -isnot [datetime] }
 }
}

function Select-StaleSub ($months) {
 begin {
  $subCutOffDate = (Get-Date).AddMonths(-$months)
 }
 process {
  $_ | Where-Object { ($_.ad.LastLogonDate -lt $subCutOffDate) -and ($_.ad.whenCreated -lt $subCutOffDate) }
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
  $_
 }
}

function Set-AeriesSTFInactive ($params) {
 begin { $sql = "UPDATE STF SET tg = 'i' WHERE id = @id" }
 process {
  $sqlVars = "id=$($_.ad.EmployeeID)"
  Write-Host ('{0},{1}' -f $MyInvocation.MyCommand.name, $_.info) -F DarkYellow
  if ($WhatIf) { return $_ }
  New-SqlOperation @params -Query $sql -Parameters $sqlVars
  $_
 }
}

function Set-ADExpireDate ([pscredential]$cred) {
 process {
  $params = @{
   Identity   = $_.ad.ObjectGUID
   DateTime   = $_.expireDate
   Credential = $cred
   Confirm    = $false
   WhatIf     = $WhatIf
  }
  Write-Host ('{0},{1},Target Date: {2}' -f $MyInvocation.MyCommand.Name, $_.info, $_.expireDate) -F Magenta
  Set-ADAccountExpiration @params
  $_
 }
}

function Set-GoogleArchiveState {
 process {
  if (!($_.googleData)) { return $_ } # Skip if no Google data
  if ($_.googleData.archived -eq $true) { return $_ } # Already archived
  Write-Host ('{0},{1},{2} - Archived On' -f $MyInvocation.MyCommand.name, $_.info, $_.ad.HomePage) -F DarkBlue
  if (!$WhatIf) {
   (& $gam update user "$($_.ad.HomePage)" archived on)*>null
   Start-Sleep -Seconds 30 # Allow time for archive state change to propagate
  }
  $_
 }
}

function Set-GoogleData {
 process {
  try { ($_.googleData = & $gam print users query "email:$($_.ad.HomePage)" fields 'archived,suspended,orgUnitPath' | ConvertFrom-Csv)*>$null }
  catch {
   Write-Host ('{0},{1},Error Getting Google Data' -f $MyInvocation.MyCommand.Name, $_.info) -F Red
  }
  $_
 }
}

function Set-PropAD ($data) {
 process {
  $id = $_.emp.EmpId
  $_.ad = $data.Where({ $_.EmployeeId -eq $id })
  if (!$_.ad) { return }
  $_
 }
}

function Set-PropEmailParams ([string[]]$manager, [string]$type) {
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

function Set-PropExchangeData ($data) {
 process {
  $upn = $_.ad.UserPrincipalName
  # if (!($data).count -eq 0) { return $_ }
  $_.exchangeData = if ($data) {
   $data.Where({ $_.UserPrincipalName -eq $upn })
  }
  else {
   $filter = "UserPrincipalName -eq '$upn'"
   Get-EXOMailbox -Filter $filter -ErrorAction SilentlyContinue
  }
  if (!($_.exchangeData)) {
   Write-Verbose ('{0},{1},Mailbox Not Found.' -f $MyInvocation.MyCommand.Name, $_.info)
  }
  $_
 }
}

function Set-PropEmp ($data) {
 process {
  $id = $_.ad.EmployeeId
  $_.emp = $data.Where({ [int]$_.empId -eq [int]$id })
  $_
 }
}

function Set-PropExpireDate ($days) {
 begin { $today = [datetime]::Today }
 process {
  $lastDay = if ($_.emp.EmploymentStatusCode -match 'S') { $today }
  elseif ($_.emp.DateTerminationLastDay -match '\d') { $_.emp.DateTerminationLastDay }
  elseif ($_.emp.DateTermination -match '\d') { $_.emp.DateTermination }
  else { $today }

  $_.expireDate = (Get-Date $lastDay).AddDays($days)
  $_
 }
}

function Set-PropInfo {
 process {
  $msg = $_.ad.EmployeeID, $_.ad.Mail, $_.ad.AccountExpirationDate, $_.ad.LastLogonDate
  $_.info = '[{0},{1},Exp:[{2}],LastLogon:[{3}]]' -f $msg
  $_
 }
}

function Set-RandomPassword ([pscredential]$cred) {
 process {
  Write-Host('{0},{1}' -f $MyInvocation.MyCommand.name, $_.info) -F DarkYellow
  $params = @{
   Identity    = $_.ad.ObjectGUID
   NewPassword = ConvertTo-SecureString -String (New-RandomPassword 20) -AsPlainText -Force
   Credential  = $cred
   Confirm     = $false
   WhatIf      = $WhatIf
  }
  Set-ADAccountPassword @params
  $_
 }
}

function Show-Object {
 begin { $i = 0 }
 process {
  if ($_) {
   Write-Verbose ($MyInvocation.MyCommand.name, $_ | Out-String)
   $i++
  }
 }
 end { Write-Verbose ('Total objects processed: {0}' -f $i) }
}

function Skip-ADDisabled {
 process {
  if ($_.ad.Enabled -eq $false) { return }
  $_
 }
}

function Suspend-GoogleAccount {
 process {
  if (!($_.googleData)) { return $_ } # Skip if no Google data
  if ($_.googleData.suspended -eq $true) { return $_ } # Already suspended
  Write-Host ('{0},{1},{2}' -f $MyInvocation.MyCommand.name, $_.info, $_.ad.HomePage) -F Blue
  if (!$WhatIf) { (& $gam update user "$($_.ad.HomePage)" suspended on)*>null }
  $_
 }
}

function Update-ADDescription ($msg, [pscredential]$cred) {
 process {
  if ($_.ad.Description -match 'Expiration Date') { return $_ }
  $newDesc = '{0} <{1} {2}>' -f $_.ad.Description, $msg, (Get-Date $_.expireDate -f 'yyyy-MM-dd')
  Write-Host ('{0},{1},[{2}]' -f $MyInvocation.MyCommand.Name, $_.info, $newDesc) -f Magenta
  $params = @{
   Identity    = $_.ad.ObjectGUID
   Description = $newDesc
   Credential  = $cred
   Confirm     = $false
   WhatIf      = $WhatIf
  }
  Set-ADUser @params
  $_
 }
}

function Update-OrgUnitAD ($ou, [pscredential]$cred) {
 process {
  if ($_.ad.DistinguishedName -match $ou ) { return $_ }
  Write-Host ('{0},{1},[{2}]' -f $MyInvocation.MyCommand.Name, $_.info, $ou) -F Magenta
  $params = @{
   Identity   = $_.ad.ObjectGUID
   TargetPath = $ou
   Confirm    = $false
   Credential = $cred
   WhatIf     = $WhatIf
  }
  Move-ADObject @params
  $_
 }
}

function Update-OrgUnitGoogle ($ou) {
 process {
  if (!$_.googleData) { return $_ }
  if ($_.googleData.orgUnitPath -match $ou) { return $_ }
  Write-Host ('{0},{1},[{2}]' -f $MyInvocation.MyCommand.Name, $_.info, $ou) -F Magenta
  if (!$WhatIf) { (& $gam update user "$($_.ad.HomePage)" org "$ou")*>null }
  $_
 }
}

# Main ==========================================================================
Import-Module Mailozaurr -Cmdlet Send-EmailMessage
Import-Module dbatools -Cmdlet Invoke-DbaQuery, Set-DbatoolsConfig, Connect-DbaInstance, Disconnect-DbaInstance
Import-Module ExchangeOnlineManagement -Cmdlet Connect-ExchangeOnline, Disconnect-ExchangeOnline, Get-EXOMailbox
Import-Module CommonScriptFunctions

$gam = 'C:\GAM7\gam.exe'

if ($WhatIf) { Show-TestRun }
Clear-SessionData

$empDBParams = @{
 Server     = $EmpServer
 Database   = $EmpDatabase
 Credential = $EmpCredential
}

$SISParams = @{
 Server     = $SISServer
 Database   = $SISDatabase
 Credential = $SISCredential
}

$adProps = 'EmployeeID', 'Mail', 'AccountExpirationDate', 'accountExpires', 'HomePage', 'info',
'memberOf', 'Description', 'msDS-ExternalDirectoryObjectId', 'LastLogonDate', 'WhenCreated'
$adData = Get-ADData -props $adProps -cred $ADCredential

Show-BlockInfo 'Set Stale Sub Expirations'
Get-EmpRetired $empDBParams (Get-Content -Path '.\sql\retired-subs.sql' -Raw) |
 Format-Object |
  Set-PropAD -data $adData |
   Select-ADWithoutExpiration |
    Skip-ADDisabled |
     Set-PropInfo |
      Select-StaleSub -months $StaleSubMonths |
       Set-PropExpireDate -days $GracePeriodDays |
        Set-ADExpireDate -cred $ADCredential |
         Update-ADDescription -msg 'Expiration Date Set' -cred $ADCredential |
          Show-Object

Show-BlockInfo 'Set Account Expirations'
Get-EmpRetired -sqlParams $empDBParams -query (Get-Content -Path '.\sql\retired-terminated.sql' -Raw) |
 Format-Object |
  Set-PropAD -data $adData |
   Select-ADWithoutExpiration |
    Skip-ADDisabled |
     Set-PropInfo |
      Set-PropExpireDate -days $GracePeriodDays |
       Set-ADExpireDate -cred $ADCredential |
        Update-ADDescription -msg 'Expiration Date Set' -cred $ADCredential |
         Show-Object

New-ExchangeOnlineSession $ExchangeOnlineCredentials

# # $exchangeData = Get-ExchangeData # Maybe not needed
$empData = Get-EmpData -sqlParams $empDBParams

$adObjsWithExpiration = $adData |
 Format-Object |
  Select-ADWithExpiration |
   Set-PropEmp -data $empData |
    Set-PropInfo

Show-BlockInfo 'Account Expiration Emails'
$adObjsWithExpiration |
 Skip-ADDisabled |
  Select-ADExpiringSoon -days $GracePeriodDays |
   Format-EmailExpiring |
    Send-ExpireAlert $EmailCredential $BccAddress |
     Show-Object

Show-BlockInfo 'Disabling Expired Accounts'
$disableAccounts = $adObjsWithExpiration | Skip-ADDisabled | Select-ADExpired
$disableAccounts |
 Set-GoogleData |
  Set-PropExchangeData -data $exchangeData |
   Convert-ToSharedMailbox |
    Clear-ExtensionAttributes -cred $ADCredential |
     Remove-ADGroupMemberships -exGroups $ProtectedGroups -cred $ADCredential |
      Disable-GoogleForwarding |
       Set-AeriesSTFInactive $SISParams |
        Remove-FromO365Groups |
         Remove-O365Forwarding |
          Suspend-GoogleAccount |
           Set-GoogleArchiveState |
            Update-OrgUnitAD -ou $OrgUnitNoLicAD -cred $ADCredential |
             Update-OrgUnitGoogle -ou $OrgUnitNoLicGoogle |
              Remove-GoogleLicense -ou $OrgUnitNoLicGoogle |
               Set-RandomPassword -cred $ADCredential |
                Disable-ADAccount -cred $ADCredential |
                 Show-Object

Show-BlockInfo 'Email disabled accounts report'
$disableAccounts |
 Set-PropEmailParams $AccountManagerEmail Disabled |
  Send-ExpireAlert $EmailCredential |
   Show-Object

Show-BlockInfo 'Remove Old Expired Accounts'
$adObjsWithExpiration |
 Select-ADExpiredOld -months $DisabledMaxMonths |
  Set-PropExchangeData -data $exchangeData |
   Remove-UGNReference $sisParams |
    Remove-GoogleAccount |
     Remove-EmpEmailWork $empDBParams |
      Disable-PreventAccidentalDelete -cred $ADCredential |
       Remove-StaffADObject -cred $ADCredential |
        Remove-SharedMailbox |
         Set-PropEmailParams $AccountManagerEmail Removed |
          Send-ExpireAlert $EmailCredential |
           Show-Object

Disconnect-ExchangeOnline -Confirm:$false
# =========================================================================================

Remove-Module tmp*

if ($WhatIf) { Show-TestRun }