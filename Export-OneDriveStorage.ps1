# Export-OneDriveStorage.ps1

#Requires -Version 7.3
<#
.SYNOPSIS
    Exports OneDrive for Business storage data to CSV.
.PARAMETER OutputFolder
    Path where the CSV report will be saved.
.EXAMPLE
    .\Export-OneDriveStorage.ps1 -OutputFolder "C:\Reports"
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
        "Microsoft.Graph.Reports"
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
    
    $Script:LogFile = Join-Path $OutputFolder "OneDriveStorage_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Write-Log "=== OneDrive Storage Export Started ===" "INFO"
    
    # Install modules
    Install-GraphModules
    
    # Clear cache and connect
    Clear-GraphCache
    Write-Log "Connecting to Microsoft Graph..." "INFO"
    Write-Log "Please authenticate when prompted..." "INFO"
    
    $Scopes = @("User.Read.All", "Reports.Read.All")
    Connect-MgGraph -Scopes $Scopes -NoWelcome
    
    $Context = Get-MgContext
    if ($Context) {
        Write-Log "Connected to Microsoft Graph successfully." "SUCCESS"
        Write-Log "Tenant: $($Context.TenantId)" "INFO"
    } else {
        throw "Failed to connect to Microsoft Graph."
    }
    
    # Get OneDrive usage report
    Write-Log "Retrieving OneDrive usage report..." "INFO"
    $TempFile = Join-Path $env:TEMP "OneDriveUsage_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    Get-MgReportOneDriveUsageAccountDetail -Period D30 -OutFile $TempFile
    
    if (-not (Test-Path $TempFile)) {
        throw "OneDrive usage report was not generated."
    }
    
    $OneDriveData = Import-Csv $TempFile
    Write-Log "Found $($OneDriveData.Count) OneDrive accounts." "INFO"
    
    $Report = @()
    
    foreach ($Item in $OneDriveData) {
        try {
            $ReportLine = [PSCustomObject]@{
                ServiceType = "OneDrive"
                DisplayName = $Item.'Owner Display Name'
                UserPrincipalName = $Item.'Owner Principal Name'
                PrimarySmtpAddress = $Item.'Owner Principal Name'
                RecipientType = "OneDriveUser"
                StorageUsedMB = if ($Item.'Storage Used (Byte)') { [Math]::Round([long]$Item.'Storage Used (Byte)' / 1MB, 2) } else { 0 }
                StorageQuotaMB = if ($Item.'Storage Allocated (Byte)') { [Math]::Round([long]$Item.'Storage Allocated (Byte)' / 1MB, 2) } else { 0 }
                ItemCount = if ($Item.'File Count') { $Item.'File Count' } else { 0 }
                LastLogonTime = if ($Item.'Last Activity Date') { $Item.'Last Activity Date' } else { "N/A" }
                Database = "OneDrive"
                AdditionalInfo = "Active Files: $(if ($Item.'Active File Count') { $Item.'Active File Count' } else { '0' })"
            }
            $Report += $ReportLine
        }
        catch {
            Write-Log "Error processing OneDrive data for $($Item.'Owner Display Name'): $_" "WARN"
        }
    }
    
    # Export results
    $OutputPath = Join-Path $OutputFolder "OneDriveStorageData_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $Report | Sort-Object StorageUsedMB -Descending | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    
    Write-Log "OneDrive storage data exported to: $OutputPath" "SUCCESS"
    Write-Log "Total OneDrive accounts processed: $($Report.Count)" "INFO"
    
    $TotalUsedGB = [Math]::Round(($Report | Measure-Object -Property StorageUsedMB -Sum).Sum / 1024, 2)
    Write-Log "Total OneDrive storage used: $TotalUsedGB GB" "INFO"
    
    # Cleanup
    Remove-Item $TempFile -ErrorAction SilentlyContinue
    Disconnect-MgGraph
    Write-Log "=== OneDrive Storage Export Completed ===" "SUCCESS"
}
catch {
    Write-Log "Fatal error: $_" "ERROR"
    try { Disconnect-MgGraph } catch { }
    exit 1
}
