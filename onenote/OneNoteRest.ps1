
$configFile = $pwd.Path + "\..\config.json"
$PSObject=Get-Content -Path $configFile | ConvertFrom-Json

# Authenticate to Microsoft Grpah
 Write-Host "Authenticating to Microsoft Graph via REST method"
 
$url = "https://login.microsoftonline.com/"+$PSObject.tenantID+"/oauth2/token"
$resource = "https://graph.microsoft.com/"
$restbody = @{
         grant_type    = 'client_credentials'
         client_id     = $PSObject.applicationID
         client_secret = $PSObject.clientKey
         resource      = $resource
}
     
 # Get the return Auth Token
$token = Invoke-RestMethod -Method POST -Uri $url -Body $restbody
     
# Set the baseurl to MS Graph-API (BETA API)
# Pack the token into a header for future API calls
$header = @{
          'Authorization' = "$($Token.token_type) $($Token.access_token)"
         'Content-type'  = "application/json"
}
 
# Build the URL for the API call
# Get a list of notebooks you have access to
$url = $PSObject.baseUrl + '/groups/' + $PSObject.groupID + '/onenote/notebooks'
$url  
# Call the REST-API
$notebooklist = Invoke-RestMethod -Method GET -headers $header -Uri $url
 
write-host $notebooklist.value