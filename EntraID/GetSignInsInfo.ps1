## This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.  
## THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
## INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  
## We grant You a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of the Sample Code, 
## provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the Sample Code is embedded; 
## (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, 
## and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, that arise or result from the use or distribution of the Sample Code.
## Please note: None of the conditions outlined in the disclaimer above will supersede the terms and conditions contained within the Premier Customer Services Description.
##
## Description 
## the PS script inventory enterprise application from EntraID
##
## Instruction 
## 1. Update $clientId to point to AAD application Id 
## 2. Update $tenantName to point to AAD tenant name
## 3. Update $thumbprint to include correct certificate thumbprint
## 4. Update $siteUrl to point to a SPO site collection
## Output: 
## 1. A CSV file with the list of enterprise applications and their sign-ins information

$scriptFullPath = $MyInvocation.MyCommand.Path
$scriptPath = Split-Path $scriptFullPath

$appInfo = Get-Content $("{0}\credential.json" -f $scriptPath) | ConvertFrom-Json
$clientId = $appInfo.ClientId
$tenantName = $appInfo.TenantName
$thumbprint = $appInfo.Thumbprint
$siteUrl = $appInfo.SiteUrl

$csvFile = $("{0}\EnterpriseApps.csv" -f $scriptPath)
$csvResultFile = $("{0}\EnterpriseAppsSignIns.csv" -f $scriptPath)
$days = -90
$csvRows = Import-Csv -Path $csvFile

$signInResults = @()

Connect-PnPOnline -Url $siteUrl `
    -ClientId $clientId `
    -Tenant $tenantName `
    -Thumbprint $thumbprint `
    -WarningAction Ignore

foreach ($row in $csvRows) {
    $accessToken = Get-PnPAccessToken -ResourceTypeName Graph
    $date = (Get-Date).AddDays($days).ToString("yyyy-MM-dd")
    # get the latest sign-in
    $url = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$top=1&`$filter=appId eq '$($row.appId)' and createdDateTime ge $($date)&sort=createdDateTime desc"
    $headers = @{
        
        "Authorization" = "Bearer $($accessToken)"
        "Content-Type"  = "application/json"
    }
    $restResult = Invoke-RestMethod -Headers $headers `
        -Uri $url `
        -Method Get `
        -Body $null
    $signInNew = [PSCustomObject]@{
        id                = $row.id
        appId             = $row.appId
        appDisplayName    = ""
        signInDateTime    = ""
        userDisplayName   = ""
        userPrincipalName = ""
        ipAddress         = ""
        location          = ""
        status            = ""
    }
    # if there is a sign-in record, get the details
    if ($restResult.value.Count -ne 0) {
        $signInNew.signInDateTime = $restResult.value[0].createdDateTime
        $signInNew.userDisplayName = $restResult.value[0].userDisplayName
        $signInNew.userPrincipalName = $restResult.value[0].userPrincipalName
        $signInNew.ipAddress = $restResult.value[0].ipAddress
        $signInNew.location = $restResult.value[0].location
        $signInNew.status = $restResult.value[0].status
    }
    else {
        write-host "No sign-in record found for $($row.appDisplayName) within the last $days days."
    }
    $signInResults += $signInNew
}

$signInResults | Export-Csv -Path $csvResultFile -NoTypeInformation

Disconnect-PnPOnline
