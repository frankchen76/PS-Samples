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
## the PS script reset all documents in a user's OneDrive site
##
## Instruction 
## 1. Update $clientId to point to AAD application Id 
## 2. Update $tenantName to point to AAD tenant name
## 3. Update $thumbprint to include correct certificate thumbprint
## 4. Update $adminSiteUrl to point to a SPO admin site collection

$scriptFullPath = $MyInvocation.MyCommand.Path
$scriptPath = Split-Path $scriptFullPath

$csvFile = $("{0}\Reset-OneDrivePermissionInheritance.csv" -f $scriptPath)
#$csvResultFile = $("{0}\Reset-OneDrivePermissionInheritance-Result.csv" -f $scriptPath)

if ($null -eq $scriptPath) {
    $csvFile = $("spo\permissions\{0}" -f $csvFile)
    #$csvResultFile = $("spo\permissions\{0}" -f $csvResultFile)
}

[array]$csvRows = Import-Csv $csvFile
$progress = 0

# get app-only credentials from JSON file
$appInfo = Get-Content $("credential-725618.json" -f $scriptPath) | ConvertFrom-Json
$clientId = $appInfo.ClientId
$tenantName = $appInfo.TenantName
$thumbprint = $appInfo.Thumbprint
$adminSiteUrl = $appInfo.AdminsiteUrl

# connecto to SPO admin site
Connect-PnPOnline -Url $adminSiteUrl `
    -ClientId $clientId `
    -Tenant $tenantName `
    -Thumbprint $thumbprint `
    -WarningAction Ignore
Write-Host $("Connected admin site '{0}' with thumbprint." -f $adminSiteUrl)

foreach ($csvRow in $csvRows) {
    # get user's onedrive url
    $userInfo = Get-PnPUserProfileProperty -Account $csvRow.UPN
    if ($null -eq $userInfo) {
        Write-Host ("User {0} does not have a OneDrive site." -f $csvRow.UPN)
        continue
    }

    # connect to user's OneDrive site
    Connect-PnPOnline -Url $userInfo.PersonalUrl `
        -ClientId $clientId `
        -Tenant $tenantName `
        -Thumbprint $thumbprint `
        -WarningAction Ignore
    Write-Host $("Connected user {0} OneDrive '{1}'." -f $csvRow.UPN, $userInfo.PersonalUrl)

    # get document library from OneDrive
    $docList = Get-PnPList -Identity "Documents" -Includes "HasUniqueRoleAssignments"

    Write-Host ("Processing list item level unique permission for list '{0}'..." -f $docList.Title)

    $allListItems = Get-PnPListItem -List $docList.Title -PageSize 1000 #-Fields "HasUniqueRoleAssignments", "Client_Title"
    $progress = 0
    $totalCount = $allListItems.Count
    foreach ($listItem in $allListItems) {
        $loadedListItem = Get-PnPListItem -List $docList.Title -Fields "HasUniqueRoleAssignments", "Client_Title", "LinkFilename" -Id $listItem.Id
        if ($loadedListItem.HasUniqueRoleAssignments -eq $true) {
            $loadedListItem.ResetRoleInheritance();
            $loadedListItem.Context.ExecuteQuery();
            Write-Host -ForegroundColor Green ("Removed item level unique permission for '{0}'." -f $loadedListItem.FieldValues["FileRef"])
        }
        # else {
        #     Write-Host ("No item level unique permission was found for '{0}'." -f $loadedListItem.FieldValues["LinkFilename"])
        # }
        $progress++

        Write-Progress -Activity "Remove list unique permissions..." `
            -Status ("Checked {0}/{1} documents" -f $progress, $totalCount) `
            -PercentComplete ($progress * 100 / $totalCount)
    }
    Write-Host ("List item level unique permission for list '{0}' total items: {1} was processed" -f $docList.Title, $allListItems.count)

}
