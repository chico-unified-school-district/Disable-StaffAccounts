function Show-BlockInfo {
 [cmdletbinding()]
 param($str)
 $textInfo = (Get-Culture).TextInfo
 $str = $textInfo.ToTitleCase($str)
 Write-Host ('=== {0} {1}' -f $str, ('=' * (50 - $str.length)))  -Fore DarkMagenta
 Write-Debug 'Proceed?'
}