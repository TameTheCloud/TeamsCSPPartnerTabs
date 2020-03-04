#Get Partner Credential
$PartnerCredential = (Get-Credential)

#Sign into CSP Partner Portal and Azure Tenant
Connect-MsolService -Credential $PartnerCredential
Connect-AzureRmAccount -Credential $PartnerCredential

#Teams CSP Application ID/Secret 
$ApplicationClientID = 'YOURSHERE'
$ClientSecret = 'YOURSHERE'

#Set Desired Teams Name under which each Company will be created
$TeamName = "My Customers"

#Create Authorization Request
$tenantID =  (Get-AzureRmTenant).Id
$graphUrl = 'https://graph.microsoft.com'
$tokenEndpoint = "https://login.microsoftonline.com/$tenantID/oauth2/token"

$tokenHeaders = @{
  "Content-Type" = "application/x-www-form-urlencoded";
}

$tokenBody = @{
  "grant_type"    = "client_credentials";
  "client_id"     = "$ApplicationClientID";
  "client_secret" = "$ClientSecret";
  "resource"      = "$graphUrl";
}


#Post request to get the access token so we can query the Microsoft Graph (valid for 1 hour)
$response = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Headers $tokenHeaders -Body $tokenBody

#Create the headers to send with the access token obtained from the above post
$queryHeaders = @{
  "Content-Type" = "application/json"
  "Authorization" = "Bearer $($response.access_token)"
}

#Collect Info about CSP and their Customers
$CSPDomain = Get-MsolDomain | Where-Object {$_.IsDefault -eq $true}
$companies = Get-MsolPartnerContract -All | Select-Object *

#Get Current User to Add as Group Owner later
$queryUrl = $graphUrl + "/v1.0/users" + "/$($PartnerCredential.UserName)"
$GroupOwner = Invoke-RestMethod -Method GET -Uri $queryUrl -Headers $queryHeaders


#Check for and if not exist, create the Initial Group (Required to then create a Team). 
#Note, the mailnickname is dervied from this and cannot contain spaces and other certain charactors, it will attempt to resolve it but may be the source of errors. 

$queryUrl = $graphUrl + "/v1.0/Groups" + "?`$filter=DisplayName+eq+'$($TeamName)'"
$initialGroup = Invoke-RestMethod -Method GET -Uri $queryUrl -Headers $queryHeaders
if ($initialGroup.value.length -eq 0){


$queryBody = @"
{
  "description": "$($TeamName)",
  "displayName": "$($TeamName)",
  "groupTypes": [
    "Unified"
  ],
  "mailEnabled": true,
  "mailNickname": "$($TeamName.Replace(' ','~'))",
  "visibility": "Private",
  "securityEnabled": false,
   "owners@odata.bind": ["https://graph.microsoft.com/v1.0/users/$($GroupOwner.id)"]
}
"@
$queryUrl = $graphUrl + "/v1.0/Groups"
$initialGroup = Invoke-RestMethod -Method POST -Uri $queryUrl -Headers $queryHeaders -Body $queryBody
}

#Add Teams to the Group
$queryUrl = $graphUrl + "/v1.0/Groups/$($initialGroup.id)/team"
$queryBody = "{}"
$InitialTeam = Invoke-RestMethod -Method PUT -Uri $queryUrl -Headers $queryHeaders -Body $queryBody

#Wait for the Team to create
start-sleep -Seconds 30

#Create the Channels and Tabs for each company
Foreach ($Company in $companies){
$queryBody = @"
{
  "displayName": "$($company.Name)",
  "description": "This channel is dedicated the for tenant with initial tenant: $($Company.DefaultDomainName)",
  "isFavoriteByDefault": true
}
"@

#$queryUrl
$queryUrl = $graphUrl + "/v1.0/teams/$($InitialTeam.id)/channels"

#$queryBody
  $TeamChannel = Invoke-RestMethod -Method POST -Uri $queryUrl -Headers $queryHeaders -Body $queryBody
    $Teamchannel
    Write-Output "Waiting for Channel to create..."
    

    #Create an Exchange Tab
    $queryBody = @"
    {
    "name": "Exch",
    "teamsAppId": "com.microsoft.teamspace.tab.web",
    "configuration": {
    "entityId": "",
    "contentUrl": "https://outlook.office365.com/ecp/?rfr=Admin_o365&exsvurl=1&delegatedOrg=$($Company.DefaultDomainName)&mkt=en-US&Realm=$($CSPDomain.Name)",
    "removeUrl": "",
    "websiteUrl": "https://outlook.office365.com/ecp/?rfr=Admin_o365&exsvurl=1&delegatedOrg=$($Company.DefaultDomainName)&mkt=en-US&Realm=$($CSPDomain.Name)"
        }
    }
"@
    $queryUrl = $graphUrl + "/beta/teams/$($InitialTeam.id)/channels/$($TeamChannel.id)/tabs"
    Invoke-RestMethod -Method POST -Uri $queryUrl -Headers $queryHeaders -Body $queryBody

    #Create an Intune Tab
    $queryBody = @"
    {
    "name": "Intune",
    "teamsAppId": "com.microsoft.teamspace.tab.web",
    "configuration": {
    "entityId": "",
    "contentUrl": "https://portal.azure.com/$($Company.DefaultDomainName)/#blade/Microsoft_Intune_DeviceSettings/ExtensionLandingBlade/overview",
    "removeUrl": "",
    "websiteUrl": "https://portal.azure.com/$($Company.DefaultDomainName)/#blade/Microsoft_Intune_DeviceSettings/ExtensionLandingBlade/overview"
        }
    }
"@
    $queryUrl = $graphUrl + "/beta/teams/$($InitialTeam.id)/channels/$($TeamChannel.id)/tabs"
    Invoke-RestMethod -Method POST -Uri $queryUrl -Headers $queryHeaders -Body $queryBody

    #Create an O365 Tab
    $queryBody = @"
    {
    "name": "O365",
    "teamsAppId": "com.microsoft.teamspace.tab.web",
    "configuration": {
    "entityId": "",
    "contentUrl": "https://portal.office.com/Partner/BeginClientSession.aspx?CTID=$($Company.TenantId)&CSDEST=o365admincenter",
    "removeUrl": "",
    "websiteUrl": "https://portal.office.com/Partner/BeginClientSession.aspx?CTID=$($Company.TenantId)&CSDEST=o365admincenter"
        }
    }
"@
    $queryUrl = $graphUrl + "/beta/teams/$($InitialTeam.id)/channels/$($TeamChannel.id)/tabs"
    Invoke-RestMethod -Method POST -Uri $queryUrl -Headers $queryHeaders -Body $queryBody

}
