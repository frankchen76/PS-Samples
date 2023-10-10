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
## the PS script inventory file versions using enumeration. 
##
## Instruction 
## Fill in the following variables: 
# $clientId: the AAD App Id of the app registration
# $tenantName: the tenant name
# $thumbprint: the thumbprint of the certificate
# 
# this script can inventory a listed user against to a tenant. 

$appInfo = Get-Content "..\credential.json" | ConvertFrom-Json
$clientId = $appInfo.ClientId
$tenantName = $appInfo.TenantName
$thumbprint = $appInfo.Thumbprint

# this can be any site collection
$TargetSiteUrl = "https://m365x725618.sharepoint.com/sites/FrankCommunication1"
$CSVFile = $("C:\Temp\SamplesLogs\InventoryFileVersions_enum_{0}.csv" -f [System.DateTime]::Now.ToString("yyyy-MM-dd_hh-mm-ss"))

# Connect-PnPOnline -Url $SiteUrl -ClientId $ACSAppId -ClientSecret $ACSAppSecret -WarningAction Ignore
Connect-PnPOnline -Url $TargetSiteUrl `
    -ClientId $clientId `
    -Tenant $tenantName `
    -Thumbprint $thumbprint `

Write-Host $("Connected to '{0}' using ClientId: {1}; Certificate thumbprint '{2}'." -f $SiteUrl, $clientId, $thumbprint)

# retrieve all document libraries which are not system, site assets, or hidden
$docLibs = Get-PnPList -Includes ("Title", "Created", "Hidden", "BaseTemplate", "IsSystemList", "IsSiteAssetsLibrary") | Where-Object { $_.BaseTemplate -eq 101 -and $_.IsSystemList -eq $false -and $_.IsSiteAssetsLibrary -eq $false }
Write-Host $("Found '{0}' document libraries." -f $docLibs.Count)

$progress = 0

# loop $docLibs and retrieve list items for each doc library
foreach ($docLib in $docLibs) {
    $Results = @()
    # retrieve all list items
    $listItems = Get-PnPListItem -List $docLib.Title -Fields "FileRef", "_UIVersionString", "_UIVersion", "File_x0020_Size" -PageSize 5000 | Where-Object { $_.FileSystemObjectType -eq "File" }
    # add list items properties to $Results
    foreach ($listItem in $listItems) {
        $Results += New-Object PSObject -Property @{
            "Id"              = $listItem.Id
            "DocLib"          = $docLib.Title
            "ObjectVersion"   = $listItem.ObjectVersion
            "UIVersionString" = $listItem["_UIVersionString"]
            "UIVersion"       = $listItem["_UIVersion"]
            "FileRef"         = $listItem["FileRef"]
            "FileSize"        = $listItem["File_x0020_Size"]
        }
    }
    #Export results to CSV
    $Results | Export-Csv -Path $CSVFile -NoTypeInformation -Append
    Write-Host $("{0} file information was appended to '{1}'." -f $Results.Count, $CSVFile)

    $progress++
    Write-Progress -Activity ("Processed doc library '{0}'" -f $docLib.Title) `
        -Status ("Processed {0}/{1} doc libraries" -f $progress, $docLibs.Count) `
        -PercentComplete ($progress * 100 / $docLibs.Count)
    
}

Disconnect-PnPOnline
Write-Host $("Disconnected '{0}'." -f $SiteUrl)
