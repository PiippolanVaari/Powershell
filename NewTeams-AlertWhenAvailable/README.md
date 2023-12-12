New version of MS Teams doesn't have the "Alert when available" option anymore so I wrote a powershell script that does the equivalent.
It uses Microsoft Graph powershell module to get users availability status and Windows SpVoice Interface or alternatively a message box to alert user when specified user is available.
