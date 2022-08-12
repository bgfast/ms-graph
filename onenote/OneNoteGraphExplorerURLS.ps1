
$configFile = $pwd.Path + "\..\config.json"

$PSObject=Get-Content -Path $configFile | ConvertFrom-Json

"********************************************************"
"*** Useful URLs for the graph api                    ***"
"********************************************************"
""
"*** Get your GUID and other user profile information ***"
"    https://graph.microsoft.com/v1.0/me/joinedTeams"

"*** Get a list of notebooks you have access to ***"
"    " +$PSObject.baseUrl + '/groups/' + $PSObject.groupID + '/onenote/notebooks'