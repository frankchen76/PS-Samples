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
## the PS script inventory document libraries' version setting based on a CSV file. The script can enumerate all sub sites
##
## Instruction 
## Fill in the following variables: 
# $clientId: the AAD App Id of the app registration
# $tenantName: the tenant name
# $thumbprint: the thumbprint of the certificate
# $csvFile: the input CSV file which contains SiteUrl column for Site Collection URL
# $csvOuptputFile: the output CSV file which contains SiteUrl, DocLib, EnableVersioning, MajorVersionLimit columns
# 
# this script removes the doc versions for site collections. 

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

function Get-DocumentLibVersion {
    param(
        [string]$ClientId,
        [string]$Thumbprint = $null,
        [string]$AppSecret = $null,
        [string]$TenantName,
        [string]$SiteUrl
    )

}

$clientId = "17f6377c-1e78-40ff-81f0-00e9876d0519"
$tenantName = "MngEnvMCAP604196.onmicrosoft.com"
$thumbprint = "2c1674918e4ccfbd791e0fd6ac6520755fb45e40"
$csvOuptputFile = [System.String]::Format("c:\temp\SamplesLogs\DocLibVersions_{0}.csv", [System.DateTime]::Now.ToString("yyyy-MM-dd"))
$csvFile = "DocLib-input.csv"

$Results = @()
$csvRows = Import-Csv -Path $csvFile

$index = 0

foreach ($row in $csvRows) {
    $siteUrl = $row."SiteUrl"

    #connect to site collection
    ConnectToSC -ClientId $clientId -Thumbprint $thumbprint -TenantName $tenantName -SiteUrl $siteUrl
    
    # Get all document libraries
    $docLibs = Get-PnPList -Includes ("Title", "Hidden", "BaseTemplate", "IsSystemList", "IsSiteAssetsLibrary") `
    | Where-Object { $_.BaseTemplate -eq 101 `
            -and $_.Hidden -eq $false `
            -and $_.IsSystemList -eq $false `
            -and $_.IsSiteAssetsLibrary -eq $false }
    
    Write-Host $("Retrieved {0} document libraries." -f $docLibs.Count)

    # enable the max version for documnet libraries
    foreach ($docLib in $docLibs) {
        $Results += [pscustomobject] @{
            SiteUrl           = $siteUrl
            DocLib            = $docLib.Title
            #DocLibUrl         = $docLib.ListUrl
            EnableVersioning  = $docLib.EnableVersioning
            MajorVersionLimit = $docLib.MajorVersionLimit
        }
    }

    # Get sub sites
    $subWebs = Get-PnPSubWeb -Recurse
    foreach ($subWeb in $subWebs) {
        ConnectToSC -ClientId $clientId -Thumbprint $thumbprint -TenantName $tenantName -SiteUrl $subWeb.Url
        $docLibs = Get-PnPList -Includes ("Title", "Hidden", "BaseTemplate", "IsSystemList", "IsSiteAssetsLibrary") `
        | Where-Object { $_.BaseTemplate -eq 101 `
                -and $_.Hidden -eq $false `
                -and $_.IsSystemList -eq $false `
                -and $_.IsSiteAssetsLibrary -eq $false }
        
        Write-Host $("Retrieved {0} document libraries." -f $docLibs.Count)

        # enable the max version for documnet libraries
        foreach ($docLib in $docLibs) {
            $Results += [pscustomobject] @{
                SiteUrl           = $subWeb.Url
                DocLib            = $docLib.Title
                #DocLibUrl         = $docLib.ListUrl
                EnableVersioning  = $docLib.EnableVersioning
                MajorVersionLimit = $docLib.MajorVersionLimit
            }
        }
        Disconnect-PnPOnline
    }

    $index++
    $percentage = [math]::Truncate($index * 100 / $csvRows.Count)
    Write-Progress -Activity "Process..." -Status "$percentage% Complete:" -PercentComplete $percentage;

}

$Results | export-csv -Path $csvOuptputFile -Append -NoTypeInformation
Write-Host $("CSV file was exported to {0}." -f $csvOuptputFile)

