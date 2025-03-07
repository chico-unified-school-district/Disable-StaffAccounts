function Create-O365PSSession {
 [cmdletbinding()]
 param(
  [System.Management.Automation.PSCredential]$Credential
 )

 Write-Host 'Creating Office365 session' -F Blue
 $o365SessionParams = @{
  ConfigurationName = 'Microsoft.Exchange'
  ConnectionUri     = 'https://outlook.office365.com/powershell-liveid/'
  Credential        = $Credential
  Authentication    = 'Basic'
  AllowRedirection  = $True
 }
 $O365Session = New-PSSession @o365SessionParams -ErrorAction Stop
 Import-PSSession -Session $O365Session -ErrorAction Stop
}