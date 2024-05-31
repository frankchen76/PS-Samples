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
## the PS script inventory ACS application created on SPO site collections
##
## Instruction 
## 1. Update $clientId to point to AAD application Id 
## 2. Update $tenantName to point to AAD tenant name
## 3. Update $thumbprint to include correct certificate thumbprint
## 4. Update $adminSiteUrl to point to SPO admin site collection
## 5. Update $csvFile to point to right CSV file location

function ConnectToSC {
    param(
        [string]$ClientId,
        [string]$Thumbprint = $null,
        [string]$AppSecret = $null,
        [string]$TenantName,
        [string]$SiteUrl

    )
    $isConnected = $false
    if ($null -ne $Thumbprint) {
        Connect-PnPOnline -Url $SiteUrl `
            -ClientId $ClientId `
            -Tenant $TenantName `
            -Thumbprint $Thumbprint `
            -WarningAction Ignore
        Write-Host $("Connected site '{0}' with thumbprint." -f $SiteUrl)
        $isConnected = $true
    }
    elseif ($null -ne $AppSecret) {
        Connect-PnPOnline -AppId $ClientId -AppSecret $AppSecret -Url $SiteUrl -WarningAction Ignore
        Write-Host $("Connected site '{0}' with ACS." -f $SiteUrl)
        $isConnected = $true
    }
    else {
        Write-Error "Please provide either a thumbprint or an app secret"
    }
    $isConnected
}
function GetACSInfo {
    param (
        [string]$ClientId,
        [string]$Thumbprint = $null,
        [string]$AppSecret = $null,
        [string]$TenantName,
        [string]$SiteUrl
    )

    $ret = @()
    # connect to each site collection
    $siteConnected = ConnectToSC -ClientId $ClientId `
        -Thumbprint $Thumbprint `
        -TenantName $TenantName `
        -SiteUrl $SiteUrl

    if ($siteConnected -eq $true) {
        $acsInfos = Get-PnPAzureACSPrincipal -Scope All
        Write-Host $("Retrieved {0} acs apps from site collection '{1}'." -f $acsInfos.Count, $SiteUrl)

        foreach ($acsInfo in $acsInfos) {
            # enumerate site collection scoped permissions
            foreach ($sc in $acsInfo.SiteCollectionScopedPermissions) {
                $ret += [pscustomobject] @{
                    SiteUrl              = $SiteUrl
                    AppId                = $acsInfo.AppId
                    RedirectUri          = $acsInfo.RedirectUri
                    Title                = $acsInfo.Title
                    AppDomains           = $acsInfo.AppDomains | Join-String -Separator ";"
                    ValidUntil           = $acsInfo.ValidUntil
                    SCSiteId             = $sc.SiteId
                    SCWebId              = $sc.WebId
                    SCListId             = $sc.ListId
                    SCRight              = $sc.Right
                    TenantProductFeature = ""
                    TenantScope          = ""
                    TenantRight          = ""
                    TenantResourceId     = ""
                }
            }

            # enumerate tenant scoped permissions
            foreach ($t in $acsInfo.TenantScopedPermissions) {
                $ret += [pscustomobject] @{
                    SiteUrl              = $site.Url
                    AppId                = $acsInfo.AppId
                    RedirectUri          = $acsInfo.RedirectUri
                    Title                = $acsInfo.Title
                    AppDomains           = $acsInfo.AppDomains | Join-String -Separator ";"
                    ValidUntil           = $acsInfo.ValidUntil
                    SCSiteId             = ""
                    SCWebId              = ""
                    SCListId             = ""
                    SCRight              = ""
                    TenantProductFeature = $t.ProductFeature
                    TenantScope          = $t.Scope
                    TenantRight          = $t.Right
                    TenantResourceId     = $t.ResourceId
                }
            }
        }    
        Disconnect-PnPOnline
        Write-Host $("Disconnected to '{0}'." -f $SiteUrl)

        $ret
    }

}
$scriptFullPath = $MyInvocation.MyCommand.Path
$scriptPath = Split-Path $scriptFullPath

$appInfo = Get-Content $("{0}\credential-725618.json" -f $scriptPath) | ConvertFrom-Json
$clientId = $appInfo.ClientId
$tenantName = $appInfo.TenantName
$thumbprint = $appInfo.Thumbprint

$adminSiteUrl = "https://m365x725618-admin.sharepoint.com"
$csvFile = $("C:\Temp\SamplesLogs\InventoryACSApps-output-{0}.csv" -f [System.DateTime]::Now.ToString("yyyy-MM-dd_hh-mm-ss"))
$Results = @()
$allSiteUrls = @()
$index = 0
$sw = [Diagnostics.Stopwatch]::StartNew()

$allSiteUrls += $adminSiteUrl

# connect to the admin site collection
$siteConnected = ConnectToSC -ClientId $clientId `
    -Thumbprint $thumbprint `
    -TenantName $tenantName `
    -SiteUrl $adminSiteUrl

if ($siteConnected -eq $true) {
    # add all site collection URL to array
    Get-PnPTenantSite | ForEach-Object {
        $allSiteUrls += $_.Url
    }
    Write-Host $("Retrieved {0} site collections'." -f $allSiteUrls.Count)
    
    Disconnect-PnPOnline
    Write-Host $("Disconnected to '{0}'." -f $adminSiteUrl)
}

foreach ($siteUrl in $allSiteUrls) {

    # get ACS Info for each site collection
    $acsResults = GetACSInfo  -ClientId $clientId `
        -Thumbprint $thumbprint `
        -TenantName $tenantName `
        -SiteUrl $siteUrl 
    
    $Results += $acsResults
    # display progress
    $index++
    $percentage = [math]::Truncate($index * 100 / $allSiteUrls.Count)
    Write-Progress -Activity "Process..." -Status "$percentage% Complete:" -PercentComplete $percentage;

}
#Export results to CSV
$Results | Export-Csv -Path $csvFile -NoTypeInformation
Write-Host $("File inventory csv was exported to '{0}'." -f $csvFile)

$sw.Stop()
Write-Host $("Process was completed within {0}." -f $sw.Elapsed)


