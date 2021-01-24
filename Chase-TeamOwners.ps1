### ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ###
### Requires an app registration with the Groups.Read.All and User.Read.All privileges ###
### ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ ###

# Connection variables

$TenantName = "domain.onmicrosoft.com"
$ClientID = "<CLIENT-ID"
$ClientSecret = "<CLIENT-SECRET>"

# Define body of authentication token request

$TokenBody = @{  
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $ClientID
    Client_Secret = $ClientSecret
}   

# Get authentication token

$tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantName/oauth2/v2.0/token" -Method POST -Body $TokenBody  

# List all group objects within the tenant

$URL = "https://graph.microsoft.com/v1.0/groups?$select=id,resourceProvisioningOptions"
$Groups = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $URL -Method Get  

# Get all groups linked to teams and add them to an array

$TeamsInfo = @()

ForEach ($Group in $Groups.value) {
    If ($Group.resourceProvisioningOptions -eq "Team") {
        $PSObject = New-Object PSObject
        $PSObject | Add-Member -NotePropertyName TeamName -NotePropertyValue $Group.displayName
        $PSObject | Add-Member -NotePropertyName TeamID -NotePropertyValue $Group.id
        $PSObject | Add-Member -NotePropertyName TeamDescription -NotePropertyValue $Group.description
        $PSObject | Add-Member -NotePropertyName TeamCreated -NotePropertyValue $Group.createdDateTime
        $TeamsInfo += $PSObject
        }
    }

# Get the owners of each team

#$URL = "https://graph.microsoft.com/v1.0/groups?$select=id,resourceProvisioningOptions"

$TeamOwners = @()

ForEach ($Team in $TeamsInfo) {
    $URL = ("https://graph.microsoft.com/v1.0/groups/" + $Team.TeamID + "/owners")
    $Owners = Invoke-RestMethod -Headers @{Authorization = "Bearer $($tokenResponse.access_token)"} -Uri $URL -Method Get
    $PSObject = New-Object PSObject
    $PSObject | Add-Member -NotePropertyName TeamName -NotePropertyValue $Team.TeamName
    $PSObject | Add-Member -NotePropertyName TeamID -NotePropertyValue $Team.TeamID
    $PSObject | Add-Member -NotePropertyName TeamDescription -NotePropertyValue $Team.TeamDescription
    $PSObject | Add-Member -NotePropertyName TeamCreated -NotePropertyValue $Team.TeamCreated
    $PSObject | Add-Member -NotePropertyName TeamOwnerCount -NotePropertyValue $Owners.value.Count

    $Count = 0
    ForEach ($Owner in $Owners.value) {
        $Count++
        $PSObject | Add-Member -NotePropertyName ("OwnerID" + $Count) -NotePropertyValue $Owner.id
        $PSObject | Add-Member -NotePropertyName ("OwnerDisplayName" + $Count) -NotePropertyValue $Owner.displayName
        $PSObject | Add-Member -NotePropertyName ("OwnerMail" + $Count) -NotePropertyValue $Owner.mail
        #$PSObject | Add-Member -NotePropertyName TeamName -NotePropertyValue $Team.TeamName
        }
    $TeamOwners += $PSObject
    }

# Send an email where there is only a single owner of the team

$SingleOwnerTeams = $TeamOwners | Where-Object {$_.TeamOwnerCount -eq 1}

$From = "user@domain.com"
$Subject = "Add owner to Team"

$TeamsTotal = $TeamsInfo.Count
$ProcessingCount = 0

ForEach ($Team in $SingleOwnerTeams) {

$ProcessingCount++

Write-Host ("Processing email " + $ProcessingCount + " of " + $UserTotal + "...") -ForegroundColor Yellow
Write-Host ("Sending mail to `"" + $Team.OwnerDisplayName1 + "`"...") -ForegroundColor Yellow

$Body = ("<font face=""calibri"" color:#1F497D; style=""font-size:11pt;"">
Dear " + $Team.OwnerDisplayName1.Split(" ")[0] + ",<br><br>

You're currently listed as the only owner of the following team:<br><br>

<style>
TABLE{border: 1px solid black; border-collapse: collapse; font-family: calibri; font-size: 11pt;}
TD{border: 1px solid black; padding: 5px; }
</style>

<table><colgroup><col><col><col></colgroup>
<tr><td><strong>Team Name</strong></td><td>" + $Team.TeamName + "
<tr><td><strong>Description</strong></td><td>" + $Team.TeamDescription + "
<tr><td><strong>Created</strong></td><td>" + ([datetime]::Parse($Team.TeamCreated)).ToString("dd/MM/yyyy HH:mm") + "
</table><br>

Please use this guide to add one or more additional owners to the team - <a href=""https://support.microsoft.com/en-us/office/go-to-guide-for-team-owners-92d238e6-0ae2-447e-af90-40b1052c4547"">https://support.microsoft.com/en-us/office/go-to-guide-for-team-owners-92d238e6-0ae2-447e-af90-40b1052c4547</a> . This will mean that others will be able to manage the team in your absence.<br><br>

Many thanks for your assistance.<br><br><br>

The Microsoft 365 Team

</font>")

$Outlook = New-Object -ComObject Outlook.Application

$Mail = $Outlook.CreateItem(0)
$Mail.To = $Team.OwnerMail1
$Mail.Subject = $Subject
$Mail.HTMLBody = $Body
$Mail.SentOnBehalfOfName = $From
$Mail.Send()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

Sleep 1

}
