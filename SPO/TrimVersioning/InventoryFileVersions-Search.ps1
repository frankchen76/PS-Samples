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
## the PS script inventory file versions. 
##
## Instruction 
## Fill in the following variables: 
# $clientId: the AAD App Id of the app registration
# $tenantName: the tenant name
# $thumbprint: the thumbprint of the certificate
# $SiteUrl: the site collection url
# $TargetSiteUrl: the site collection url to be inventoried
# $CSVFile: the csv file to store the inventory result
# $CSVFileByFileType: the csv file to store the inventory result by file type
# $BaselineSize: the minimum file size to filter.
# 
# this script can inventory a listed user against to a tenant. 

$appInfo = Get-Content "..\credential.json" | ConvertFrom-Json
$clientId = $appInfo.ClientId
$tenantName = $appInfo.TenantName
$thumbprint = $appInfo.Thumbprint


# this can be any site collection
$SiteUrl = "https://m365x725618.sharepoint.com"
$TargetSiteUrl = "https://m365x725618.sharepoint.com/sites/FrankCommunication1"
$CSVFile = $("C:\Temp\SamplesLogs\InventoryFileVersions_{0}.csv" -f [System.DateTime]::Now.ToString("yyyy-MM-dd_hh-mm-ss"))
$CSVFileByFileType = $("C:\Temp\SamplesLogs\InventoryFileVersions-filetype_{0}.csv" -f [System.DateTime]::Now.ToString("yyyy-MM-dd_hh-mm-ss"))
# 100MB
# $BaselineSize = 100 * 1024 * 1024
$BaselineSize = 0

Connect-PnPOnline -Url $siteUrl `
    -ClientId $clientId `
    -Tenant $tenantName `
    -Thumbprint $thumbprint `

Write-Host $("Connected to '{0}' using ClientId: {1}; Certificate thumbprint '{2}'." -f $SiteUrl, $clientId, $thumbprint)

$SearchQuery = $("contentclass:STS_ListItem_DocumentLibrary AND IsContainer:false AND IsDocument:true AND Size>{0} AND path:""{1}/*""" -f $BaselineSize, $TargetSiteUrl)
$SelectProperties = @("Path", "Url", "Title", "Author", "AuthorOWSUSER", "ModifiedById", "FileExtension", "FileType", "Created", "LastModifiedTime", "SPWebUrl", "Size", "UIVersionStringOWSTEXT")
$SortList = @{ Size = "descending" }

#Run Search Query    
Write-Host $("Execute query to '{0}'..." -f $SiteUrl)
$SearchResults = Submit-PnPSearchQuery -Query $SearchQuery `
    -SelectProperties $SelectProperties `
    -SortList $SortList `
    -All

Write-Host $("Query was completed, {0} records were retrieved." -f $SearchResults.TotalRows)

$Results = @()
ForEach ($ResultRow in $SearchResults.ResultRows) {    
    $Results += [pscustomobject] @{
        Path          = $ResultRow["Path"]
        Url           = $ResultRow["Url"]
        Title         = $ResultRow["Title"]
        Author        = $ResultRow["Author"]
        AuthorOWSUSER = $ResultRow["AuthorOWSUSER"]
        ModifiedById  = $ResultRow["ModifiedById"]
        FileExtension = $ResultRow["FileExtension"]
        FileType      = $ResultRow["FileType"]
        Created       = $ResultRow["Created"]
        SPWebUrl      = $ResultRow["SPWebUrl"]
        LastModified  = $ResultRow["LastModifiedTime"]
        Size          = $ResultRow["Size"]
        MaxVersion    = $ResultRow["UIVersionStringOWSTEXT"]
    }
}

#Export results to CSV
$Results | Export-Csv -Path $CSVFile -NoTypeInformation
Write-Host $("File inventory csv was exported to '{0}'." -f $CSVFile)

# export report based on file type
$ResultsByFileType = @()
$groupByFileTypes = $Results | Group-Object -Property FileType
foreach ($groupByFileType in $groupByFileTypes) {
    $totalSize = 0
    $groupByFileType.Group | ForEach-Object { $totalSize += $_.Size }
    $ResultsByFileType += [pscustomobject] @{
        FileType = $groupByFileType.Name
        Count    = $groupByFileType.Count
        Size     = $totalSize
    }
}
$ResultsByFileType | Export-Csv -Path $CSVFileByFileType -NoTypeInformation
Write-Host $("File size by file type inventory csv was exported to '{0}'." -f $CSVFileByFileType)

Disconnect-PnPOnline
Write-Host $("Disconnected '{0}'." -f $SiteUrl)
