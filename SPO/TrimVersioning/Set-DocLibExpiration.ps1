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
## this script apply version expiration policy to a site collection.  
##
## Instruction 
## Fill in the following variables or create a credential.json file in the parent folder:
# $clientId: the AAD App Id of the app registration
# $tenantName: the tenant name
# $thumbprint: the thumbprint of the certificate
# $siteUrl: the site collection url
# $daysExpired: the number of days to keep the version
# $versionsToKeep: the number of versions to keep


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

function New-PnPFileVersionBatchDeleteJob {
    <#
    .SYNOPSIS
        Starts to delete file versions in batches.  

    .DESCRIPTION
        This cmdlet allows users to specify a scope and a cutoff time, where 
        all the file versions satisfies the following are permanently deleted:

            1) created before the specified time, and
            2) exisit within the specified scope.
        
        Scope can either by a Site or a List (must be document library).
        
    .PARAMETER Site 
        A Microsoft.SharePoint.Client.Site object. By using this parameter, 
        the job is scoped to the site level. 

    .PARAMETER Library 
        A Microsoft.SharePoint.Client.List object. By using this parameter, 
        the job is scoped to the list level. 
        
        Note that this must be a document library.

    .PARAMETER DeleteBeforeDays 
        The number of days that specifies the cut off time for deletion. 
        I.e. file versions created before this number of days ago will be 
        permanently deleted.

        Note that this number cannot be smaller than 30, and cannot be 
        before the year 2023. 

    .EXAMPLE
        # Example 1. Scope: Site, Delete Versions Created Before: 50 days ago
        Connect-PnPOnline -Url "https://contoso.sharepoint.com" -UseWebLogin
        $Site = Get-PnPSite
        New-PnPFileVersionBatchDeleteJob -Site $Site -DeleteBeforeDays 50

    .EXAMPLE
        # Example 2. Scope: List, Delete Versions Created Before: 30 days ago
        Connect-PnPOnline -Url "https://contoso.sharepoint.com" -UseWebLogin
        $Library = Get-PnPList Documents
        New-PnPFileVersionBatchDeleteJob -Library $Library -DeleteBeforeDays 30
    #>
    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'High'
    )]
    Param(
        [Parameter(Mandatory,
            ParameterSetName = 'SiteScope')]
        [Microsoft.SharePoint.Client.Site] $Site,

        [Parameter(Mandatory,
            ParameterSetName = 'LibraryScope')]
        [Microsoft.SharePoint.Client.List] $Library,

        [Parameter(ParameterSetName = 'SiteScope')]
        [Parameter(ParameterSetName = 'LibraryScope')]
        [Parameter(Mandatory)]
        [int] $DeleteBeforeDays
    )

    if (-not $PSCmdlet.ShouldProcess($PSCmdlet.ParameterSetName, "permanent version deletion")) {
        Write-Host "Cancelled. No versions will be deleted."
        return 
    }

    $Ctx = Get-PnPContext
    if ($PsCmdlet.ParameterSetName -eq "SiteScope") {
        $ClientObject = $Site
    }
    else {
        $ClientObject = $Library
    }

    $ClientObject.StartDeleteFileVersions($DeleteBeforeDays)
    $Ctx.ExecuteQuery()
    Write-Host "Success. Versions specified will be deleted in the upcoming days."
}

function Remove-PnPFileVersionBatchDeleteJob {
    <#
    .SYNOPSIS
        Cancels any further progress of file version batch deletion. 

    .DESCRIPTION
        Cancels any further progress of file version batch deletion. Note that
        any version that are already deleted will not be reverted.

    .PARAMETER Site 
        A Microsoft.SharePoint.Client.Site object that has the ongoing job.

    .PARAMETER Library 
        A Microsoft.SharePoint.Client.List object that has the ongoing job.

    .EXAMPLE
        # Example 1. Stopping onging site-scoped version deletion job. 
        Connect-PnPOnline -Url "https://contoso.sharepoint.com" -UseWebLogin
        $Site = Get-PnPSite
        Remove-PnPFileVersionBatchDeleteJob -Site $Site

    .EXAMPLE
        # Example 2. Stopping onging list version deletion job. 
        Connect-PnPOnline -Url "https://contoso.sharepoint.com" -UseWebLogin
        $Library = Get-PnPList Documents
        Remove-PnPFileVersionBatchDeleteJob -Library $Library
    #>
    [CmdletBinding(
        SupportsShouldProcess,
        ConfirmImpact = 'High'
    )]
    Param(
        [Parameter(Mandatory,
            ParameterSetName = 'SiteScope')]
        [Microsoft.SharePoint.Client.Site] $Site,

        [Parameter(Mandatory,
            ParameterSetName = 'LibraryScope')]
        [Microsoft.SharePoint.Client.List] $Library
    )

    if (-not $PSCmdlet.ShouldProcess($PSCmdlet.ParameterSetName, "stop further version deletion batches")) {
        Write-Host "Did not receive confirmation to stop deletion. Continuing to delete specified versions."
        return 
    }

    $Ctx = Get-PnPContext
    if ($PsCmdlet.ParameterSetName -eq "SiteScope") {
        $ClientObject = $Site
    }
    else {
        $ClientObject = $Library
    }

    $ClientObject.CancelDeleteFileVersions()
    $Ctx.ExecuteQuery()
    Write-Host "Future deletion is successfully stopped."
}

function ApplyExpirationToSC {
    param(
        [string]$ClientId,
        [string]$Thumbprint = $null,
        [string]$AppSecret = $null,
        [string]$TenantName,
        [string]$SiteUrl,
        [int]$VersionsToKeep,
        [int]$DaysExpired
    )
    $siteConnected = ConnectToSC -ClientId $ClientId `
        -Thumbprint $Thumbprint `
        -TenantName $TenantName `
        -SiteUrl $SiteUrl

    if ($siteConnected -eq $true) {
        # apply version limit to SC
        Set-PnPSite -Identity $SiteUrl `
            -EnableAutoExpirationVersionTrim $true `
            -ExpireVersionsAfterDays $DaysExpired `
        
        # Get all document libraries
        $docLibs = Get-PnPList -Includes ("Title", "Hidden", "BaseTemplate", "IsSystemList", "IsSiteAssetsLibrary") `
        | Where-Object { $_.BaseTemplate -eq 101 `
                -and $_.Hidden -eq $false `
                -and $_.IsSystemList -eq $false `
                -and $_.IsSiteAssetsLibrary -eq $false }

        # enable the max version for documnet libraries
        foreach ($docLib in $docLibs) {
            # apply the max version number when Versioning is enabled
            if ($docLib.EnableVersioning) {
                # configure the max version for lib if $maxLibVersion is not equal to 0. 
                if ($maxLibVersion -ne 0) {
                    $docLib.MajorVersionLimit = $VersionsToKeep
                    $docLib.Update()
                    $context.ExecuteQuery()
                    Write-Host $("Applied max version number '{0}' to document library '{1}'." -f $versionsToKeep, $docLib.Title)
                }
                else {
                    Write-Host $("Skipped applying max version number to document library '{0}' as maxLibVersion is 0." -f $versionsToKeep, $docLib.Title)
                }
        
            }
            else {
                Write-Host $("Skiped doc lib '{0}' as it didn't enable versioning" -f $docLib.title)
            }
        
            $index++
            $percentage = [math]::Truncate($index * 100 / $docLibs.Count)
            Write-Progress -Activity "Process..." -Status "$percentage% Complete:" -PercentComplete $percentage;
    
        }

        # initiate the automatic version deletion job
        $currentSite = Get-PnPSite
        New-PnPFileVersionBatchDeleteJob -Site $currentSite -DeleteBeforeDays $DaysExpired

        Disconnect-PnPOnline
        Write-Host $("Disconnected site '{0}'." -f $SiteUrl)
    }

}

$appInfo = Get-Content "..\credential.json" | ConvertFrom-Json
$clientId = $appInfo.ClientId
$tenantName = $appInfo.TenantName
$thumbprint = $appInfo.Thumbprint

$siteUrl = "https://m365x725618.sharepoint.com/sites/DocVersioning"
# max days to keep the version
$daysExpired = 180
# max version number to keep
$versionsToKeep = 50

ApplyExpirationToSC -ClientId $clientId `
    -Thumbprint $thumbprint `
    -TenantName $tenantName `
    -SiteUrl $siteUrl `
    -VersionsToKeep $versionsToKeep `
    -DaysExpired $daysExpired
