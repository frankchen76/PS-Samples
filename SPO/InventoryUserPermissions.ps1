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
## this script can inventory a listed user against to a tenant.
##
## Instruction 
## 1. create InventoryUserPermissions.csv including one column "UPN" and list all users to be inventoried 
## 2. execute the ps and enter the SPO tenant admin user credential

$scriptFullPath = $MyInvocation.MyCommand.Path
$scriptPath = Split-Path $scriptFullPath

$csvFile = $("{0}\InventoryUserPermissions.csv" -f $scriptPath)
$csvResultFile = $("{0}\InventoryUserPermissions-Result.csv" -f $scriptPath)
[array]$csvRows = Import-Csv $csvFile
$result = @()
$progress = 0

if ($csvRows.Count -gt 0) {
    Write-Host $("Read {0} rows from '{1}'." -f $csvRows.count, $csvFile)

    $spoAdminUrl = "https://m365x725618-admin.sharepoint.com"
    Connect-SPOService -Url $spoAdminUrl
    Write-Host $("Connected to '{0}'" -f $spoAdminUrl)

    $allSites = Get-SPOSite -Limit ALL
    Write-Host $("Found {0} site collections." -f $allSites.Count)

    foreach ($site in $allSites) {
        
        foreach ($row in $csvRows) {
            try {
                $user = $null
                $user = Get-SPOUser -Site $site.Url -LoginName $row."UPN" -ErrorAction SilentlyContinue
            }
            catch {}
            
            if ($null -ne $user) {
                
                $userRow = @{
                    UPN    = $row."UPN"
                    Url    = $site.Url
                    Groups = $user.Groups -join "," 
                }
                $result += New-Object PSObject -Property $userRow
                Write-Host $("User {0} can access site collections {1}." -f $row."UPN", $site.Url)
            }
        }
        $progress++

        Write-Progress -Activity ("Inventory user permission from site collection '{0}'..." -f $site.Url) `
            -Status ("Checked {0}/{1} site collections" -f $progress, $allSites.count) `
            -PercentComplete ($progress * 100 / $allSites.count)
    }
}

$result | Export-Csv $csvResultFile
Write-Host $("Result were export to '{0}'" -f $csvResultFile)

Disconnect-SPOService