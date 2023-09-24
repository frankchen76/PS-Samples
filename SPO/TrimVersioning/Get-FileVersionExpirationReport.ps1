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
## this script return file version expiration report for a file.  
##
## Instruction 
## Fill in the following variables or create a credential.json file in the parent folder:
# $clientId: the AAD App Id of the app registration
# $tenantName: the tenant name
# $thumbprint: the thumbprint of the certificate
 

function Get-PnPFileVersionExpirationReport {
    <#
    .SYNOPSIS

        Retrieve a file version expiration report.

    .DESCRIPTION

        Retrieves the version history of a file with its estimated snapshot time (FileVersion.SnapshotDate)
        and expiration time under the automatic expiration policy (FileVersion.ExpirationDate). Note it does
        not include the file's current version.

    .PARAMETER Url

        The server relative URL to the file.

    .PARAMETER Automatic

        Simulate automatic expiration version policy.

    .PARAMETER FixedNumberOfDays

        Simulate a time-based expiration version policy.

    .EXAMPLE

        Connect-PnPOnline -Url https://contoso.sharepoint.com -UseWebLogin
        $Report = Get-PnPFileVersionExpirationReport -Url "/Shared Documents/MyFile.txt" -Automatic
        $Report | Select-Object VersionLabel, SnapshotDate, ExpirationDate

    .EXAMPLE

        Connect-PnPOnline -Url https://contoso.sharepoint.com -UseWebLogin
        $Report = Get-PnPFileVersionExpirationReport -Url "/Shared Documents/MyFile.txt" -FixedNumberOfDays 120
        $Report | Select-Object VersionLabel, SnapshotDate, ExpirationDate

    #>

    [CmdletBinding(DefaultParameterSetName = 'Auto')]
    Param(
        [Parameter(ParameterSetName = 'Auto')]
        [Parameter(ParameterSetName = 'TimeBased')]
        [Parameter(Mandatory)]
        [string] $Url,

        [Parameter(ParameterSetName = 'Auto')]
        [switch] $Automatic,

        [Parameter(ParameterSetName = 'TimeBased')]
        [int] $FixedNumberOfDays
    )
    if (-not ([System.Management.Automation.PSTypeName]'VersionUtils.PropertyUtils').Type) {
        Add-Type `
            -IgnoreWarnings `
            -ReferencedAssemblies @( 
            [System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client").Location,
            [System.Reflection.Assembly]::Load("Microsoft.SharePoint.Client.Runtime").Location,
            [System.Reflection.Assembly]::Load("PnP.Framework").Location,
            "netstandard",
            "System.Runtime",
            "System.Linq.Expressions") `
            -TypeDefinition @"
                using Microsoft.SharePoint.Client;
                using PnP.Framework;

                namespace VersionUtils
                {
                    public class PropertyUtils
                    {   
                        public static FileVersionCollection LoadFileVersionExpirationReport(PnPClientContext context, File file)
                        {
                            context.Load(file, f => f.VersionExpirationReport.IncludeWithDefaultProperties(v => v.SnapshotDate, v => v.ExpirationDate));
                            context.ExecuteQuery();
                            return file.VersionExpirationReport;
                        }
                    }
                } 
"@
    }

    $Ctx = Get-PnPContext
    $File = Get-PnPFile -Url $Url

    $VersionExpirationReport = [VersionUtils.PropertyUtils]::LoadFileVersionExpirationReport($Ctx, $File);

    if ($PSCmdlet.ParameterSetName -eq 'Auto') {
        return $VersionExpirationReport
    }

    return $VersionExpirationReport |
    ForEach-Object {
        Add-Member -PassThru `
            -Force `
            -InputObject $PSItem `
            -MemberType ScriptProperty `
            -Name ExpirationDate `
            -Value {
            if ($null -eq $this.SnapshotDate) { return $null }
            return [DateTime]::ParseExact($this.SnapshotDate, "o", $null).AddDays($FixedNumberOfDays).ToString("o")
        }.GetNewClosure()
    }
}
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

$appInfo = Get-Content "..\credential.json" | ConvertFrom-Json
$clientId = $appInfo.ClientId
$tenantName = $appInfo.TenantName
$thumbprint = $appInfo.Thumbprint

$siteUrl = "https://m365x725618.sharepoint.com/sites/DocVersioning"
$fileUrl = "/Shared%20Documents/VersioningTest.docx"

ConnectToSC -ClientId $clientId `
    -TenantName $tenantName `
    -Thumbprint $thumbprint `
    -SiteUrl $siteUrl

$report = Get-PnPFileVersionExpirationReport -Url $fileUrl `
    -Automatic

$report | Select-Object VersionLabel, SnapshotDate, ExpirationDate