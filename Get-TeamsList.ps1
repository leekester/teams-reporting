# Input Parameters

$clientId = "<CLIENT-ID>"
$clientSecret = "<CLIENT-SECRET>"
$tenantName = "tenantname.onmicrosoft.com"
$resource = "https://graph.microsoft.com/"
$URL = "https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"

$tokenBody = @{  
    Grant_Type    = "client_credentials"  
    Scope         = "https://graph.microsoft.com/.default"  
    Client_Id     = $clientId  
    Client_Secret = $clientSecret  
}   
  
$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $tokenBody  
$result = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $URL -Method Get  
$TeamsInfo = ($result | select-object Value).Value | Select-Object id, displayName, visibility

# Get owner for each team

$Owners = @()

ForEach ($Team in $TeamsInfo) {
    Write-Host $Team.displayName
    $GroupID = $Team.ID
    $URL = $null
    $URL = ('https://graph.microsoft.com/beta/groups/' + $GroupID + '/owners?$select=id')
    $Results = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $URL -Method Get
    ForEach ($Result in $Results) {
        Write-Host $Results.Count
        Write-Host $Result.value.id
        }
    
    }

#$URL = "https://graph.microsoft.com/beta/groups/?`$filter=resourceProvisioningOptions/Any(x:x eq 'Team')"
