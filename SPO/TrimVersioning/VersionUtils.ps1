function Get-PnPFileVersionExpirationReport 
{
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
    if (-not ([System.Management.Automation.PSTypeName]'VersionUtils.PropertyUtils').Type)
    {
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

    if ($PSCmdlet.ParameterSetName -eq 'Auto')
    {
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

function New-PnPFileVersionBatchDeleteJob 
{
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

    if (-not $PSCmdlet.ShouldProcess($PSCmdlet.ParameterSetName, "permanent version deletion")){
        Write-Host "Cancelled. No versions will be deleted."
        return 
    }

    $Ctx = Get-PnPContext
    if ($PsCmdlet.ParameterSetName -eq "SiteScope")
    {
        $ClientObject = $Site
    }
    else {
        $ClientObject = $Library
    }

    $ClientObject.StartDeleteFileVersions($DeleteBeforeDays)
    $Ctx.ExecuteQuery()
    Write-Host "Success. Versions specified will be deleted in the upcoming days."
}

function Remove-PnPFileVersionBatchDeleteJob 
{
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

    if (-not $PSCmdlet.ShouldProcess($PSCmdlet.ParameterSetName, "stop further version deletion batches")){
        Write-Host "Did not receive confirmation to stop deletion. Continuing to delete specified versions."
        return 
    }

    $Ctx = Get-PnPContext
    if ($PsCmdlet.ParameterSetName -eq "SiteScope")
    {
        $ClientObject = $Site
    }
    else {
        $ClientObject = $Library
    }

    $ClientObject.CancelDeleteFileVersions()
    $Ctx.ExecuteQuery()
    Write-Host "Future deletion is successfully stopped."
}

function Test-AssemblyBuildVersion
{
    param (
        [Parameter(Mandatory=$true)]
        [String]$AssemblyName,

        [Parameter(Mandatory=$true)]
        [Int32]$BuildVersion
    )

    $assembly = $null

    try
    {
        $assembly = [System.Reflection.Assembly]::Load($AssemblyName)
    }
    catch
    {
        return $false
    }

    $ci = Get-ChildItem $assembly.Location
    return $ci.VersionInfo.ProductBuildPart -ge $BuildVersion
}

function Get-PnPFileVersions
{
    <#
    .SYNOPSIS

        Retrieve a file versions with SnapshotDate and ExpirationDate

    .DESCRIPTION

        Retrieves the version history of a file with its SnapshotDate and ExpirationDate

    .PARAMETER Url

        The server relative URL to the file.

    .EXAMPLE

        Connect-PnPOnline -Url https://contoso.sharepoint.com -UseWebLogin
        $FileVersions = Get-PnPFileVersions -Url "/Shared Documents/MyFile.txt"
        $FileVersions | Select-Object VersionLabel, SnapshotDate, ExpirationDate

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string] $Url
    )
    if (-not ([System.Management.Automation.PSTypeName]'VersionUtils.PropertyUtils2').Type)
    {
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
                    public class PropertyUtils2
                    {   
                        public static FileVersionCollection LoadFileVersions(PnPClientContext context, File file)
                        {
                            context.Load(file, f => f.Versions.IncludeWithDefaultProperties(v => v.SnapshotDate, v => v.ExpirationDate));
                            context.ExecuteQuery();
                            return file.Versions;
                        }
                    }
                } 
"@
    }

    $Ctx = Get-PnPContext
    $File = Get-PnPFile -Url $Url

    $fileVersions = [VersionUtils.PropertyUtils2]::LoadFileVersions($Ctx, $File);

    return $fileVersions;
}
