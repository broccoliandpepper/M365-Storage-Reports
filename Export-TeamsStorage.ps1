# Export-TeamsStorage.ps1

#Requires -Version 7.3
<#
.SYNOPSIS
    Exports Microsoft Teams storage data to CSV.
.PARAMETER OutputFolder
    Path where the CSV report will be saved.
.EXAMPLE
    .\Export-TeamsStorage.ps1 -OutputFolder "C:\Reports"
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$OutputFolder
)

# Global variables
$Script:LogFile = $null
$Script:StartTime = Get-Date

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp [$Level] $Message"
    
    if ($Script:LogFile) {
        $LogEntry | Out-File -FilePath $Script:LogFile -Append -Encoding UTF8
    }
    
    $Color = switch ($Level) {
        "ERROR" { "Red" }
        "WARN" { "Yellow" }
        "SUCCESS" { "Green" }
        default { "White" }
    }
    Write-Host $LogEntry -ForegroundColor $Color
}

function Install-GraphModules {
    $RequiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Reports",
        "Microsoft.Graph.Sites"
    )
    
    foreach ($ModuleName in $RequiredModules) {
        try {
            $Module = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
            
            if (-not $Module) {
                Write-Log "Installing $ModuleName module..." "INFO"
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber
                Write-Log "$ModuleName module installed successfully." "SUCCESS"
            } else {
                Write-Log "$ModuleName module (v$($Module.Version)) is available." "SUCCESS"
            }
        }
        catch {
            Write-Log "Failed to install $ModuleName module: $_" "ERROR"
            throw
        }
    }
}

function Clear-GraphCache {
    try {
        Write-Log "Clearing Microsoft Graph cache..." "INFO"
        $TokenCachePath = "$env:USERPROFILE\.mg"
        if (Test-Path $TokenCachePath) {
            Remove-Item $TokenCachePath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Graph cache cleared." "SUCCESS"
        }
    }
    catch {
        Write-Log "Error clearing cache: $_" "WARN"
    }
}

# Initialize
try {
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }
    
    $Script:LogFile = Join-Path $OutputFolder "TeamsStorage_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Write-Log "=== Teams Storage Export Started ===" "INFO"
    
    # Install modules
    Install-GraphModules
    
    # Clear cache and connect
    Clear-GraphCache
    Write-Log "Connecting to Microsoft Graph..." "INFO"
    Write-Log "Please authenticate when prompted..." "INFO"
    
    $Scopes = @("User.Read.All", "Reports.Read.All", "Sites.Read.All")
    Connect-MgGraph -Scopes $Scopes -NoWelcome
    
    $Context = Get-MgContext
    if ($Context) {
        Write-Log "Connected to Microsoft Graph successfully." "SUCCESS"
        Write-Log "Tenant: $($Context.TenantId)" "INFO"
    } else {
        throw "Failed to connect to Microsoft Graph."
    }
    
    # Get SharePoint sites usage report
    Write-Log "Retrieving SharePoint sites usage report..." "INFO"
    $TempFile = Join-Path $env:TEMP "SharePointUsage_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    Get-MgReportSharePointSiteUsageDetail -Period D30 -OutFile $TempFile
    
    if (-not (Test-Path $TempFile)) {
        throw "SharePoint usage report was not generated."
    }
    
    $SharePointData = Import-Csv $TempFile
    Write-Log "Found $($SharePointData.Count) SharePoint sites." "INFO"
    
    # Filter for Teams sites
    $TeamsData = $SharePointData | Where-Object { 
        $_.'Site Type' -eq 'Group' -or $_.'Root Web Template' -like '*TEAM*' -or $_.'Site URL' -like '*/teams/*'
    }
    
    Write-Log "Filtered to $($TeamsData.Count) Teams-related sites." "INFO"
    
    $Report = @()
    
    foreach ($Item in $TeamsData) {
        try {
            # Extract team name from URL
            $TeamName = if ($Item.'Site URL') {
                $UrlParts = $Item.'Site URL' -split '/'
                if ($UrlParts.Count -gt 0) { $UrlParts[-1] } else { "Unknown" }
            } else { "Unknown" }
            
            $ReportLine = [PSCustomObject]@{
                ServiceType = "Teams"
                DisplayName = $TeamName
                UserPrincipalName = "N/A"
                PrimarySmtpAddress = "N/A"
                RecipientType = "TeamsChannels"
                StorageUsedMB = if ($Item.'Storage Used (Byte)') { [Math]::Round([long]$Item.'Storage Used (Byte)' / 1MB, 2) } else { 0 }
                StorageQuotaMB = if ($Item.'Storage Allocated (Byte)') { [Math]::Round([long]$Item.'Storage Allocated (Byte)' / 1MB, 2) } else { 0 }
                ItemCount = if ($Item.'File Count') { $Item.'File Count' } else { 0 }
                LastLogonTime = if ($Item.'Last Activity Date') { $Item.'Last Activity Date' } else { "N/A" }
                Database = "SharePointTeams"
                AdditionalInfo = "Site URL: $($Item.'Site URL')"
            }
            $Report += $ReportLine
        }
        catch {
            Write-Log "Error processing Teams site data for $($Item.'Site URL'): $_" "WARN"
        }
    }
    
    # Export results
    $OutputPath = Join-Path $OutputFolder "TeamsStorageData_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $Report | Sort-Object StorageUsedMB -Descending | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    
    Write-Log "Teams storage data exported to: $OutputPath" "SUCCESS"
    Write-Log "Total Teams sites processed: $($Report.Count)" "INFO"
    
    $TotalUsedGB = [Math]::Round(($Report | Measure-Object -Property StorageUsedMB -Sum).Sum / 1024, 2)
    Write-Log "Total Teams storage used: $TotalUsedGB GB" "INFO"
    
    # Cleanup
    Remove-Item $TempFile -ErrorAction SilentlyContinue
    Disconnect-MgGraph
    Write-Log "=== Teams Storage Export Completed ===" "SUCCESS"
}
catch {
    Write-Log "Fatal error: $_" "ERROR"
    try { Disconnect-MgGraph } catch { }
    exit 1
}
