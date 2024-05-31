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
## 4. Update $adminSiteUrl to point to a SPO admin site collection
## Output: 
## 1. A CSV file with the list of site collection where SGName has permissions to access

$scriptFullPath = $MyInvocation.MyCommand.Path
$scriptPath = Split-Path $scriptFullPath

$csvFile = $("{0}\InventorySG.csv" -f $scriptPath)
$csvResultFile = $("{0}\InventorySG-Result.csv" -f $scriptPath)

[array]$csvRows = Import-Csv $csvFile
$permissionResults = @()
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

# get all site collections
$allSites = Get-PnPTenantSite | Where-Object { ($_.Title -notlike "" -and $_.Template -notlike "*Redirect*" -and $_.Url -notlike "*my.sharepoint.com*") }

Disconnect-PnPOnline
Write-Host $("Disconnected admin site '{0}' and retrieved {1} site collections." -f $adminSiteUrl, $allSites.Count)

foreach ($site in $allSites) {
    # connecto to SPO site
    Connect-PnPOnline -Url $site.Url `
        -ClientId $clientId `
        -Tenant $tenantName `
        -Thumbprint $thumbprint `
        -WarningAction Ignore
        
    Write-Host $("Connected to site '{0}'." -f $site.Url)

    $wrs = Get-PnPWeb -Includes RoleAssignments
    ForEach ($ra in $wrs.RoleAssignments) {
        #Get the Permission Levels assigned and Member
        Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings, Member
 
        #Get the Permission Levels assigned
        $pls = ($ra.RoleDefinitionBindings | Select-Object -ExpandProperty Name | Where-Object { ($_ -ne "Limited Access") -and ($_ -ne "Web-Only Limited Access") } ) -join ","
        If ($pls.Length -eq 0 -or $ra.Member.Title -clike '*Limited Access System Group*') { Continue }

        $SitePermissionType = $ra.Member.PrincipalType
        
        If ($SitePermissionType -eq "SharePointGroup") {
            $GroupMembers = Get-PnPGroupMember -Identity $ra.Member.Title
            ForEach ($GroupMember in $GroupMembers) {
                $groupExist = $csvRows | Where-Object { $_.SGName -eq $GroupMember.Title }
                if ($null -ne $groupExist) {
                    $permissionResults += New-Object PSObject -Property ([ordered]@{
                            SiteName         = $site.Title
                            SiteURL          = $site.url
                            From             = $ra.Member.Title
                            Group            = $GroupMember.Title
                            PermissionLevels = $pls
                        })
                    Write-Host $("Site: '{0}', SGName: {1}, Permission: {2}." -f $site.Url, $GroupMember.Title, $pls)    
                }
            }
        }
        Else {
            $groupExist = $csvRows | Where-Object { $_.SGName -eq $ra.Member.Title }
            if ($null -ne $groupExist) {
                $permissionResults += New-Object PSObject -Property ([ordered]@{
                        SiteName         = $site.Title
                        SiteURL          = $site.url
                        From             = "Direct Permission"
                        Group            = $ra.Member.Title
                        PermissionLevels = $pls
                    })
                Write-Host $("Site: '{0}', SGName: {1}, Permission: {2}." -f $site.Url, $ra.Member.Title, $pls)    

            }
        }
    }

    # update progress
    $progress++
    $percentage = [math]::Truncate($progress * 100 / $allSites.Count)
    Write-Progress -Activity "Process..." -Status "$percentage% Complete:" -PercentComplete $percentage;

    Disconnect-PnPOnline
    Write-Host $("Disconnected to site '{0}'." -f $site.Url)

}
# export the results to CSV
$permissionResults | Export-Csv -Path $csvResultFile -NoTypeInformation
