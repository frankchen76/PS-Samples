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
## the PS script remove a user from all site collections
##
## Instruction 
## 1. Update $clientId to point to AAD application Id 
## 2. Update $tenantName to point to AAD tenant name
## 3. Update $thumbprint to include correct certificate thumbprint
## 4. Update $siteUrl to point to SPO root site collection

function Parse-JWTtoken {
 
    [cmdletbinding()]
    param([Parameter(Mandatory = $true)][string]$token)
 
    #Validate as per https://tools.ietf.org/html/rfc7519
    #Access and ID tokens are fine, Refresh tokens will not work
    if (!$token.Contains(".") -or !$token.StartsWith("eyJ")) { Write-Error "Invalid token" -ErrorAction Stop }
 
    #Header
    $tokenheader = $token.Split(".")[0].Replace('-', '+').Replace('_', '/')
    #Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
    while ($tokenheader.Length % 4) { Write-Verbose "Invalid length for a Base-64 char array or string, adding ="; $tokenheader += "=" }
    Write-Verbose "Base64 encoded (padded) header:"
    Write-Verbose $tokenheader
    #Convert from Base64 encoded string to PSObject all at once
    Write-Verbose "Decoded header:"
    [System.Text.Encoding]::ASCII.GetString([system.convert]::FromBase64String($tokenheader)) | ConvertFrom-Json | fl | Out-Default
 
    #Payload
    $tokenPayload = $token.Split(".")[1].Replace('-', '+').Replace('_', '/')
    #Fix padding as needed, keep adding "=" until string length modulus 4 reaches 0
    while ($tokenPayload.Length % 4) { Write-Verbose "Invalid length for a Base-64 char array or string, adding ="; $tokenPayload += "=" }
    Write-Verbose "Base64 encoded (padded) payoad:"
    Write-Verbose $tokenPayload
    #Convert to Byte array
    $tokenByteArray = [System.Convert]::FromBase64String($tokenPayload)
    #Convert to string array
    $tokenArray = [System.Text.Encoding]::ASCII.GetString($tokenByteArray)
    Write-Verbose "Decoded array in JSON format:"
    Write-Verbose $tokenArray
    #Convert from JSON to PSObject
    $tokobj = $tokenArray | ConvertFrom-Json
    Write-Verbose "Decoded Payload:"
    
    return $tokobj
}

$accessToken = ""
function Get-ValidAccessToken {
    if ($accessToken -eq "") {
        $accessToken = Get-PnPAccessToken -ResourceTypeName SharePoint
    }
    else {
        $token = Parse-JWTtoken -token $accessToken
        $exp = $token.exp
        $now = [int][double]::Parse((Get-Date -UFormat %s))
        if ($exp -lt $now) {
            $accessToken = Get-PnPAccessToken
        }
    }
    $accessToken
}

function Delete-UserFromSC {
    param(
        [Parameter(Mandatory = $true)][string]$siteUrl,
        [Parameter(Mandatory = $true)][string]$upn
    )
    # get the latest access token
    $token = Get-ValidAccessToken
    $headers = @{

        "Authorization" = "Bearer $($token)"
        "Content-Type"  = "application/json"
        "Accept"        = "application/json;odata=verbose"
    }
    $url = $("{0}/_api/web/SiteUsers/GetByEmail(@email)?@email='{1}'" -f $siteUrl, $upn)

    try {
        $result = Invoke-RestMethod -Headers $headers `
            -Uri $url `
            -Method Delete `
            -Body $null
        
        Write-Host $("User {0} was deleted from site collection {1}" -f $upn, $siteUrl)
    }
    catch {
        # $errorCode = $_.Exception.Response.StatusCode.value__
        # $errorCode = $_.Exception.Response.StatusCode
        # $isSuccess = $_.Exception.Response.IsSuccessStatusCode
        Write-Host $("User {0} was not deleted from site collection {1} because of {2}" -f $upn, $siteUrl, $_.Exception.Response.StatusCode)
    }
}

$scriptFullPath = $MyInvocation.MyCommand.Path
$scriptPath = Split-Path $scriptFullPath

$appInfo = Get-Content $("{0}\credential.json" -f $scriptPath) | ConvertFrom-Json
$clientId = $appInfo.ClientId
$tenantName = $appInfo.TenantName
$thumbprint = $appInfo.Thumbprint


$siteUrl = "https://m365x725618.sharepoint.com"
$upn = "Alexw1@m365x725618.onmicrosoft.com"

# connecto to SPO admin site
Connect-PnPOnline -Url $siteUrl `
    -ClientId $clientId `
    -Tenant $tenantName `
    -Thumbprint $thumbprint `
    -WarningAction Ignore
Write-Host $("Connected site '{0}' with thumbprint." -f $siteUrl)

# get all site collections
$siteCollections = Get-PnPTenantSite

# Delete-UserFromSC -siteUrl $siteCollections[0].Url -upn $upn

# iterate through all site collections
$index = 0
foreach ($siteCollection in $siteCollections) {
    Delete-UserFromSC -siteUrl $siteCollection.Url -upn $upn
    $index++

    # sleep 100ms to avoid throttling
    if ($index % 10 -eq 0) {
        Write-Host "Sleeping for 100 ms"
        Start-Sleep -Milliseconds 100
    }
}
