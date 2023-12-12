<#
 $user is a string that matches some part of the users DisplayName
#>
[CmdletBinding()]
Param(
  [Parameter(Mandatory=$True,Position=1)]
   [string]$user
)
$null = get-command Connect-MgGraph -ErrorAction SilentlyContinue
if (!$?) {
	Write-Host "Error: module Microsoft.Graph not installed"
	exit 1
}
$voice = New-Object -ComObject Sapi.spvoice
Connect-MgGraph -Scopes "Presence.Read.All","User.ReadBasic.All" -NoWelcome
$id = (Get-MgUser -Property Id,DisplayName | ?{$_.DisplayName -match $user}).id
while (!$free)  {
	$avail = (Get-MgUserPresence -UserId $id).Availability
	if ($avail -eq 'Available') {
		$free =$true
	} else {
		Write-Host -NoNewline .
		Start-Sleep 15
	}
}
$voice.speak("Hey, user $user is available now")
if (!$?) {
	Add-Type -AssemblyName PresentationFramework
	$Message = "User $user is available now"
	[System.Windows.MessageBox]::Show($Message)
}
