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
