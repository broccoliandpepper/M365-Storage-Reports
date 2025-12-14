#Requires -Version 7.3
<#
.SYNOPSIS
    Exports Microsoft 365 storage consumption data (Teams, Mailboxes, OneDrive) to CSV.
.DESCRIPTION
    Collects storage usage data from Microsoft Teams (SharePoint sites), Exchange Online mailboxes, 
    and OneDrive for Business accounts and exports consolidated data to CSV files.
    Features robust error handling, automatic module installation with proper privileges,
    and comprehensive logging.
.PARAMETER OutputFolder
    Path where the CSV reports will be saved.
.PARAMETER IncludeTeams
    Include Teams storage data in the report (default: $true).
.PARAMETER IncludeMailboxes
    Include mailbox size data in the report (default: $true).
.PARAMETER IncludeOneDrive
    Include OneDrive storage data in the report (default: $true).
.PARAMETER MaxRetries
    Maximum number of retry attempts for failed operations (default: 3).
.EXAMPLE
    .\Export-M365StorageReport.ps1 -OutputFolder "C:\StorageReports"
.EXAMPLE
    .\Export-M365StorageReport.ps1 -OutputFolder "C:\Reports" -IncludeOneDrive:$false -IncludeTeams:$false
.NOTES
    Author: Enhanced Storage Report Script
    Version: 2.0
    Requires: PowerShell 7.3+, ExchangeOnlineManagement, Microsoft.Graph modules
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Specify the output folder path")]
    [string]$OutputFolder,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeTeams = $true,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeMailboxes = $true,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeOneDrive = $true,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3
)

# Global variables
$Script:LogFile = $null
$Script:StartTime = Get-Date

#region Helper Functions

function Write-Log {
    <#
    .SYNOPSIS
    Enhanced logging function with multiple severity levels and console output.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory = $false)]
        [switch]$NoConsole
    )
    
    try {
        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $LogEntry = "$Timestamp [$Level] $Message"
        
        # Write to log file
        if ($Script:LogFile) {
            $LogEntry | Out-File -FilePath $Script:LogFile -Append -Encoding UTF8
        }
        
        # Write to console with color coding
        if (-not $NoConsole) {
            $Color = switch ($Level) {
                "ERROR" { "Red" }
                "WARN" { "Yellow" }
                "SUCCESS" { "Green" }
                "DEBUG" { "Gray" }
                default { "White" }
            }
            Write-Host $LogEntry -ForegroundColor $Color
        }
        
        # Write to PowerShell streams
        switch ($Level) {
            "ERROR" { Write-Error $Message -ErrorAction SilentlyContinue }
            "WARN" { Write-Warning $Message }
            "DEBUG" { Write-Debug $Message }
            "INFO" { Write-Verbose $Message }
        }
    }
    catch {
        Write-Warning "Failed to write log entry: $_"
    }
}

function Test-PowerShellVersion {
    <#
    .SYNOPSIS
    Verifies PowerShell version compatibility.
    #>
    $RequiredVersion = [Version]"7.3.0"
    $CurrentVersion = $PSVersionTable.PSVersion
    
    if ($CurrentVersion -lt $RequiredVersion) {
        Write-Log "PowerShell version $CurrentVersion detected. Version $RequiredVersion or higher is required." "ERROR"
        throw "Unsupported PowerShell version."
    }
    
    Write-Log "PowerShell version $CurrentVersion is compatible." "SUCCESS"
}

function Test-AdminPrivileges {
    <#
    .SYNOPSIS
    Verifies that the script is running with administrator privileges.
    Returns $true if admin, $false if not (but doesn't throw).
    #>
    try {
        $CurrentPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
        $IsAdmin = $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        
        if ($IsAdmin) {
            Write-Log "Running with Administrator privileges." "SUCCESS"
            return $true
        } else {
            Write-Log "Not running as Administrator. Will use CurrentUser scope for module installation." "WARN"
            return $false
        }
    }
    catch {
        Write-Log "Failed to verify administrator privileges: $_" "WARN"
        return $false
    }
}

function Install-RequiredModule {
    <#
    .SYNOPSIS
    Safely installs a PowerShell module with proper error handling and verification.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory = $false)]
        [string]$MinimumVersion,
        
        [Parameter(Mandatory = $false)]
        [int]$RetryCount = 0,
        
        [Parameter(Mandatory = $false)]
        [bool]$IsAdmin = $false
    )
    
    try {
        Write-Log "Checking module: $ModuleName" "INFO"
        
        # Check if module is already installed
        $InstalledModule = Get-Module -ListAvailable -Name $ModuleName | 
                          Sort-Object Version -Descending | 
                          Select-Object -First 1
        
        if ($InstalledModule) {
            if ($MinimumVersion) {
                $MinVersion = [Version]$MinimumVersion
                if ($InstalledModule.Version -ge $MinVersion) {
                    Write-Log "Module $ModuleName (v$($InstalledModule.Version)) is already installed and meets requirements." "SUCCESS"
                    return $true
                }
                else {
                    Write-Log "Module $ModuleName (v$($InstalledModule.Version)) is installed but below minimum version $MinimumVersion." "WARN"
                }
            }
            else {
                Write-Log "Module $ModuleName (v$($InstalledModule.Version)) is already installed." "SUCCESS"
                return $true
            }
        }
        
        # Install or update module
        Write-Log "Installing/updating module: $ModuleName" "INFO"
        
        # Set TLS 1.2 for secure connection to PowerShell Gallery
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        # Determine installation scope based on admin privileges
        $Scope = if ($IsAdmin) { "AllUsers" } else { "CurrentUser" }
        
        $InstallParams = @{
            Name = $ModuleName
            Scope = $Scope
            Force = $true
            AllowClobber = $true
            Repository = "PSGallery"
            ErrorAction = "Stop"
        }
        
        if ($MinimumVersion) {
            $InstallParams.MinimumVersion = $MinimumVersion
        }
        
        Install-Module @InstallParams
        
        # Verify installation
        $VerifyModule = Get-Module -ListAvailable -Name $ModuleName | 
                       Sort-Object Version -Descending | 
                       Select-Object -First 1
        
        if ($VerifyModule) {
            Write-Log "Module $ModuleName (v$($VerifyModule.Version)) installed successfully." "SUCCESS"
            return $true
        }
        else {
            throw "Module installation verification failed."
        }
    }
    catch {
        Write-Log "Failed to install module $ModuleName : $_" "ERROR"
        
        if ($RetryCount -lt $MaxRetries) {
            Write-Log "Retrying module installation (attempt $($RetryCount + 1)/$MaxRetries)" "WARN"
            Start-Sleep -Seconds (2 * ($RetryCount + 1))
            return Install-RequiredModule -ModuleName $ModuleName -MinimumVersion $MinimumVersion -RetryCount ($RetryCount + 1) -IsAdmin $IsAdmin
        }
        
        throw "Failed to install required module $ModuleName after $MaxRetries attempts."
    }
}

function Initialize-RequiredModules {
    <#
    .SYNOPSIS
    Ensures all required modules are installed and available.
    #>
    param(
        [Parameter(Mandatory = $false)]
        [bool]$IsAdmin = $false
    )
    
    $RequiredModules = @()
    
    Write-Log "Determining required modules based on selected services..." "INFO"
    
    if ($IncludeMailboxes) {
        $RequiredModules += @{ Name = "ExchangeOnlineManagement"; MinVersion = "3.0.0" }
        Write-Log "Added ExchangeOnlineManagement module (Mailboxes enabled)" "DEBUG"
    }
    
    if ($IncludeTeams -or $IncludeOneDrive) {
        $RequiredModules += @{ Name = "Microsoft.Graph.Authentication"; MinVersion = "1.0.0" }
        $RequiredModules += @{ Name = "Microsoft.Graph.Reports"; MinVersion = "1.0.0" }
        $RequiredModules += @{ Name = "Microsoft.Graph.Users"; MinVersion = "1.0.0" }
        $RequiredModules += @{ Name = "Microsoft.Graph.Sites"; MinVersion = "1.0.0" }
        Write-Log "Added Microsoft Graph modules (Teams/OneDrive enabled)" "DEBUG"
    }
    
    Write-Log "Verifying $($RequiredModules.Count) required modules..." "INFO"
    
    foreach ($Module in $RequiredModules) {
        try {
            $Success = Install-RequiredModule -ModuleName $Module.Name -MinimumVersion $Module.MinVersion -IsAdmin $IsAdmin
            if (-not $Success) {
                throw "Module verification failed."
            }
        }
        catch {
            Write-Log "Critical error with module $($Module.Name): $_" "ERROR"
            throw "Cannot continue without required modules."
        }
    }
    
    Write-Log "All required modules are available." "SUCCESS"
}

function Convert-BytesToMB {
    <#
    .SYNOPSIS
    Converts various size formats to MB.
    #>
    param(
        [Parameter(Mandatory = $true)]
        $SizeValue
    )
    
    if ($SizeValue -eq $null -or $SizeValue -eq "" -or $SizeValue -eq "Unlimited") {
        return 0
    }
    
    # Handle string format like "1.5 GB (1,610,612,736 bytes)"
    if ($SizeValue -is [string]) {
        # Extract bytes from parentheses
        $BytesMatch = [regex]::Match($SizeValue, '\(([0-9,]+)\s+bytes\)')
        if ($BytesMatch.Success) {
            $BytesString = $BytesMatch.Groups[1].Value -replace ',', ''
            $Bytes = [long]$BytesString
            return [Math]::Round($Bytes / 1MB, 2)
        }
        
        # Handle direct size format like "1.5 GB"
        $SizeMatch = [regex]::Match($SizeValue, '([0-9.]+)\s*(MB|GB|KB|TB)')
        if ($SizeMatch.Success) {
            $Number = [double]$SizeMatch.Groups[1].Value
            $Unit = $SizeMatch.Groups[2].Value.ToUpper()
            
            switch ($Unit) {
                "KB" { return [Math]::Round($Number / 1024, 2) }
                "MB" { return [Math]::Round($Number, 2) }
                "GB" { return [Math]::Round($Number * 1024, 2) }
                "TB" { return [Math]::Round($Number * 1024 * 1024, 2) }
            }
        }
    }
    
    # Handle numeric values (assume bytes)
    if ($SizeValue -is [long] -or $SizeValue -is [int]) {
        return [Math]::Round($SizeValue / 1MB, 2)
    }
    
    return 0
}

function Connect-ExchangeOnlineWithRetry {
    <#
    .SYNOPSIS
    Establishes connection to Exchange Online with retry logic.
    #>
    param([int]$RetryCount = 0)
    
    try {
        Write-Log "Connecting to Exchange Online..." "INFO"
        
        # Import module explicitly
        Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop
        
        # Connect with modern authentication
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
        
        # Verify connection
        $TestCommand = Get-OrganizationConfig -ErrorAction Stop | Select-Object -First 1
        if ($TestCommand) {
            Write-Log "Successfully connected to Exchange Online." "SUCCESS"
            return $true
        }
        else {
            throw "Connection verification failed."
        }
    }
    catch {
        Write-Log "Exchange Online connection attempt failed: $_" "ERROR"
        
        if ($RetryCount -lt $MaxRetries) {
            Write-Log "Retrying Exchange connection (attempt $($RetryCount + 1)/$MaxRetries)" "WARN"
            Start-Sleep -Seconds (5 * ($RetryCount + 1))
            return Connect-ExchangeOnlineWithRetry -RetryCount ($RetryCount + 1)
        }
        
        throw "Failed to connect to Exchange Online after $MaxRetries attempts."
    }
}

function Connect-MgGraphWithRetry {
    param([int]$RetryCount = 0)
    
    try {
        Write-Log "Connecting to Microsoft Graph (attempt $($RetryCount + 1))..." "INFO"
        
        # Clear cache on first attempt
        if ($RetryCount -eq 0) {
            Reset-GraphTokenCache
        }
        
        # Define required scopes
        $Scopes = @("User.Read.All", "Reports.Read.All", "Sites.Read.All")
        
        # Try different authentication methods based on retry count
        switch ($RetryCount) {
            0 {
                # First attempt: Standard interactive
                Write-Log "Attempting standard interactive authentication..." "INFO"
                Connect-MgGraph -Scopes $Scopes -NoWelcome -ContextScope Process -ErrorAction Stop
            }
            1 {
                # Second attempt: Device code with explicit browser
                Write-Log "Attempting device code authentication..." "INFO"
                Write-Log "A browser window should open or you'll see a device code..." "INFO"
                Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome -ContextScope Process -ErrorAction Stop
            }
            2 {
                # Third attempt: Force browser authentication
                Write-Log "Attempting browser authentication..." "INFO"
                Connect-MgGraph -Scopes $Scopes -NoWelcome -ContextScope Process -UseDeviceAuthentication -ErrorAction Stop
            }
        }
        
        # Verify connection
        Start-Sleep -Seconds 3
        $Context = Get-MgContext -ErrorAction Stop
        if ($Context -and $Context.TenantId) {
            Write-Log "Successfully connected to Microsoft Graph." "SUCCESS"
            Write-Log "Connected tenant: $($Context.TenantId)" "DEBUG"
            Write-Log "Authentication method: $($Context.AuthType)" "DEBUG"
            return $true
        }
        else {
            throw "Connection verification failed - no valid context."
        }
    }
    catch {
        Write-Log "Microsoft Graph connection attempt $($RetryCount + 1) failed: $_" "ERROR"
        
        if ($RetryCount -lt ($MaxRetries - 1)) {
            Write-Log "Retrying Graph connection in 10 seconds..." "WARN"
            Start-Sleep -Seconds 10
            return Connect-MgGraphWithRetry -RetryCount ($RetryCount + 1)
        }
        
        throw "Failed to connect to Microsoft Graph after $MaxRetries attempts."
    }
}


function Export-MailboxStorageData {
    <#
    .SYNOPSIS
    Exports mailbox storage data with comprehensive error handling.
    #>
    Write-Log "Collecting mailbox storage data..." "INFO"
    
    try {
        # Connect to Exchange Online
        Connect-ExchangeOnlineWithRetry
        
        # Get all mailboxes and their statistics
        Write-Log "Retrieving mailbox list..." "INFO"
        $Mailboxes = Get-Mailbox -ResultSize Unlimited -ErrorAction Stop
        Write-Log "Found $($Mailboxes.Count) mailboxes to process." "INFO"
        
        $Report = [System.Collections.Generic.List[Object]]::new()
        
        $i = 0
        foreach ($Mailbox in $Mailboxes) {
            $i++
            $PercentComplete = [Math]::Round(($i / $Mailboxes.Count) * 100, 1)
            Write-Progress -Activity "Processing Mailboxes" -Status "Processing $($Mailbox.DisplayName) ($i/$($Mailboxes.Count))" -PercentComplete $PercentComplete
            
            try {
                $Stats = Get-MailboxStatistics -Identity $Mailbox.DistinguishedName -WarningAction SilentlyContinue -ErrorAction Stop
                
                $ReportLine = [PSCustomObject]@{
                    ServiceType = "Mailbox"
                    DisplayName = $Mailbox.DisplayName
                    UserPrincipalName = $Mailbox.UserPrincipalName
                    PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                    RecipientType = $Mailbox.RecipientTypeDetails
                    StorageUsedMB = if ($Stats.TotalItemSize) { Convert-BytesToMB -SizeValue $Stats.TotalItemSize.ToString() } else { 0 }
                    StorageQuotaMB = if ($Mailbox.ProhibitSendReceiveQuota -ne "Unlimited") { Convert-BytesToMB -SizeValue $Mailbox.ProhibitSendReceiveQuota.ToString() } else { 0 }
                    ItemCount = if ($Stats.ItemCount) { $Stats.ItemCount } else { 0 }
                    LastLogonTime = $Stats.LastLogonTime
                    Database = $Mailbox.Database
                    AdditionalInfo = "Archive: $(if ($Mailbox.ArchiveName) { 'Yes' } else { 'No' })"
                }
                $Report.Add($ReportLine)
            }
            catch {
                Write-Log "Error processing mailbox $($Mailbox.DisplayName): $_" "WARN"
                
                # Add entry with error info
                $ErrorLine = [PSCustomObject]@{
                    ServiceType = "Mailbox"
                    DisplayName = $Mailbox.DisplayName
                    UserPrincipalName = $Mailbox.UserPrincipalName
                    PrimarySmtpAddress = $Mailbox.PrimarySmtpAddress
                    RecipientType = $Mailbox.RecipientTypeDetails
                    StorageUsedMB = 0
                    StorageQuotaMB = 0
                    ItemCount = 0
                    LastLogonTime = "N/A"
                    Database = $Mailbox.Database
                    AdditionalInfo = "Error retrieving statistics"
                }
                $Report.Add($ErrorLine)
            }
        }
        
        Write-Progress -Activity "Processing Mailboxes" -Completed
        
        $OutputPath = Join-Path $OutputFolder "MailboxStorageData_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Report | Sort-Object StorageUsedMB -Descending | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        
        Write-Log "Mailbox storage data exported to: $OutputPath" "SUCCESS"
        Write-Log "Total mailboxes processed: $($Report.Count)" "INFO"
        
        # Calculate and log summary stats
        $TotalUsedMB = ($Report | Measure-Object -Property StorageUsedMB -Sum).Sum
        $TotalQuotaMB = ($Report | Where-Object { $_.StorageQuotaMB -gt 0 } | Measure-Object -Property StorageQuotaMB -Sum).Sum
        Write-Log "Mailbox summary: $([Math]::Round($TotalUsedMB / 1024, 2)) GB used, $([Math]::Round($TotalQuotaMB / 1024, 2)) GB quota" "INFO"
        
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        return $Report
    }
    catch {
        Write-Log "Error collecting mailbox data: $_" "ERROR"
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch { }
        return $null
    }
}

function Reset-GraphTokenCache {
    <#
    .SYNOPSIS
    Clears Microsoft Graph token cache to resolve authentication issues.
    #>
    try {
        Write-Log "Clearing Microsoft Graph token cache..." "INFO"
        
        # Disconnect any existing sessions
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        } catch { }
        
        # Clear token cache
        $TokenCachePath = "$env:USERPROFILE\.mg"
        if (Test-Path $TokenCachePath) {
            Remove-Item $TokenCachePath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Token cache cleared successfully." "SUCCESS"
        }
        
        # Clear additional credential caches
        $CredCachePath = "$env:LOCALAPPDATA\.IdentityService"
        if (Test-Path $CredCachePath) {
            Remove-Item $CredCachePath -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        Start-Sleep -Seconds 2
    }
    catch {
        Write-Log "Error clearing token cache: $_" "WARN"
    }
}

function Export-OneDriveStorageData {
    <#
    .SYNOPSIS
    Exports OneDrive storage data with comprehensive error handling.
    #>
    Write-Log "Collecting OneDrive storage data in separate PowerShell session..." "INFO"
    
        try {
        # Script block pour exécution dans une session propre
        $ScriptBlock = {
            param($OutputFolder, $LogFunction)
            
            try {
                # Import uniquement les modules Graph nécessaires
                Import-Module Microsoft.Graph.Authentication -Force
                Import-Module Microsoft.Graph.Reports -Force
                
                # Connexion Graph uniquement
                $Scopes = @("User.Read.All", "Reports.Read.All", "Sites.Read.All")
                Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome
                
                # Récupérer les données OneDrive
                $TempFile = Join-Path $env:TEMP "OneDriveUsage_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                Get-MgReportOneDriveUsageAccountDetail -Period D30 -OutFile $TempFile
                
                if (Test-Path $TempFile) {
                    $OneDriveData = Import-Csv $TempFile
                    $Report = @()
                    
                    foreach ($Item in $OneDriveData) {
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
                    
                    # Nettoyer et retourner les données
                    Remove-Item $TempFile -ErrorAction SilentlyContinue
                    Disconnect-MgGraph
                    return $Report
                }
                
                throw "OneDrive usage report file was not created."
            }
            catch {
                throw "OneDrive data collection failed: $_"
            }
        }
        
        # Exécuter dans une session PowerShell propre
        $Job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $OutputFolder
        $Report = Receive-Job -Job $Job -Wait
        Remove-Job -Job $Job
        
        if ($Report -and $Report.Count -gt 0) {
            $OutputPath = Join-Path $OutputFolder "OneDriveStorageData_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            $Report | Sort-Object StorageUsedMB -Descending | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            
            Write-Log "OneDrive storage data exported to: $OutputPath" "SUCCESS"
            Write-Log "Total OneDrive accounts processed: $($Report.Count)" "INFO"
            
            $TotalUsedMB = ($Report | Measure-Object -Property StorageUsedMB -Sum).Sum
            Write-Log "OneDrive summary: $([Math]::Round($TotalUsedMB / 1024, 2)) GB used" "INFO"
            
            return $Report
        } else {
            throw "No OneDrive data was returned from the job."
        }
    }
    catch {
        Write-Log "Error collecting OneDrive data: $_" "ERROR"
        return $null
    }
}

function Export-TeamsStorageData {
    <#
    .SYNOPSIS
    Exports Teams storage data with comprehensive error handling.
    #>
    Write-Log "Collecting Teams storage data..." "INFO"
    
    try {
        # Get Teams SharePoint sites usage
        Write-Log "Retrieving SharePoint sites usage report..." "INFO"
        $TempFile = Join-Path $env:TEMP "SharePointUsage_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        Get-MgReportSharePointSiteUsageDetail -Period D30 -OutFile $TempFile -ErrorAction Stop
        
        if (-not (Test-Path $TempFile)) {
            throw "SharePoint usage report file was not created."
        }
        
        $SharePointData = Import-Csv $TempFile -ErrorAction Stop
        Write-Log "Found $($SharePointData.Count) SharePoint sites." "INFO"
        
        $Report = [System.Collections.Generic.List[Object]]::new()
        
        # Filter for Teams sites (sites that have "Group" in the template or are associated with Teams)
        $TeamsData = $SharePointData | Where-Object { 
            $_.'Site Type' -eq 'Group' -or $_.'Root Web Template' -like '*TEAM*' -or $_.'Site URL' -like '*/teams/*'
        }
        
        Write-Log "Filtered to $($TeamsData.Count) Teams-related sites." "INFO"
        
        foreach ($Item in $TeamsData) {
            try {
                # Extract team name from URL (better method)
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
                $Report.Add($ReportLine)
            }
            catch {
                Write-Log "Error processing Teams site data for $($Item.'Site URL'): $_" "WARN"
            }
        }
        
        $OutputPath = Join-Path $OutputFolder "TeamsStorageData_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $Report | Sort-Object StorageUsedMB -Descending | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        
        Write-Log "Teams storage data exported to: $OutputPath" "SUCCESS"
        Write-Log "Total Teams sites processed: $($Report.Count)" "INFO"
        
        # Calculate and log summary stats
        $TotalUsedMB = ($Report | Measure-Object -Property StorageUsedMB -Sum).Sum
        Write-Log "Teams summary: $([Math]::Round($TotalUsedMB / 1024, 2)) GB used" "INFO"
        
        # Clean up temp file
        Remove-Item $TempFile -ErrorAction SilentlyContinue
        
        return $Report
    }
    catch {
        Write-Log "Error collecting Teams data: $_" "ERROR"
        return $null
    }
}

function Create-ConsolidatedReport {
    <#
    .SYNOPSIS
    Creates a consolidated storage consumption report from all services.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$AllData
    )
    
    Write-Log "Creating consolidated storage consumption report..." "INFO"
    
    try {
        # Flatten all data into a single collection
        $ConsolidatedData = [System.Collections.Generic.List[Object]]::new()
        
        foreach ($DataSet in $AllData) {
            if ($DataSet -and $DataSet.Count -gt 0) {
                foreach ($Item in $DataSet) {
                    $ConsolidatedData.Add($Item)
                }
            }
        }
        
        if ($ConsolidatedData.Count -eq 0) {
            Write-Log "No data available for consolidated report." "WARN"
            return
        }
        
        # Create summary statistics
        $MailboxData = $ConsolidatedData | Where-Object { $_.ServiceType -eq "Mailbox" }
        $OneDriveData = $ConsolidatedData | Where-Object { $_.ServiceType -eq "OneDrive" }
        $TeamsData = $ConsolidatedData | Where-Object { $_.ServiceType -eq "Teams" }
        
        $Summary = @{
            TotalMailboxes = $MailboxData.Count
            TotalOneDriveAccounts = $OneDriveData.Count
            TotalTeamsSites = $TeamsData.Count
            TotalStorageUsedMB = ($ConsolidatedData | Measure-Object -Property StorageUsedMB -Sum).Sum
            TotalStorageQuotaMB = ($ConsolidatedData | Where-Object { $_.StorageQuotaMB -gt 0 } | Measure-Object -Property StorageQuotaMB -Sum).Sum
            MailboxStorageUsedMB = ($MailboxData | Measure-Object -Property StorageUsedMB -Sum).Sum
            OneDriveStorageUsedMB = ($OneDriveData | Measure-Object -Property StorageUsedMB -Sum).Sum
            TeamsStorageUsedMB = ($TeamsData | Measure-Object -Property StorageUsedMB -Sum).Sum
        }
        
        # Export consolidated data
        $ConsolidatedPath = Join-Path $OutputFolder "ConsumedDataStorage_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $ConsolidatedData | Sort-Object StorageUsedMB -Descending | Export-Csv -Path $ConsolidatedPath -NoTypeInformation -Encoding UTF8
        
        # Create summary report
        $SummaryPath = Join-Path $OutputFolder "StorageSummary_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        $SummaryReport = [PSCustomObject]@{
            ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            TotalServices = $ConsolidatedData.Count
            TotalMailboxes = $Summary.TotalMailboxes
            TotalOneDriveAccounts = $Summary.TotalOneDriveAccounts
            TotalTeamsSites = $Summary.TotalTeamsSites
            TotalStorageUsedMB = [Math]::Round($Summary.TotalStorageUsedMB, 2)
            TotalStorageUsedGB = [Math]::Round($Summary.TotalStorageUsedMB / 1024, 2)
            TotalStorageQuotaMB = [Math]::Round($Summary.TotalStorageQuotaMB, 2)
            TotalStorageQuotaGB = [Math]::Round($Summary.TotalStorageQuotaMB / 1024, 2)
            StorageUtilizationPercent = if ($Summary.TotalStorageQuotaMB -gt 0) { [Math]::Round(($Summary.TotalStorageUsedMB / $Summary.TotalStorageQuotaMB) * 100, 2) } else { "N/A" }
            MailboxStorageUsedGB = [Math]::Round($Summary.MailboxStorageUsedMB / 1024, 2)
            OneDriveStorageUsedGB = [Math]::Round($Summary.OneDriveStorageUsedMB / 1024, 2)
            TeamsStorageUsedGB = [Math]::Round($Summary.TeamsStorageUsedMB / 1024, 2)
        }
        
        $SummaryReport | Export-Csv -Path $SummaryPath -NoTypeInformation -Encoding UTF8
        
        Write-Log "Consolidated report exported to: $ConsolidatedPath" "SUCCESS"
        Write-Log "Summary report exported to: $SummaryPath" "SUCCESS"
        Write-Log "Total storage used across all services: $([Math]::Round($Summary.TotalStorageUsedMB / 1024, 2)) GB" "INFO"
        Write-Log "Total services/accounts processed: $($ConsolidatedData.Count)" "INFO"
        
        # Log breakdown by service
        Write-Log "Storage breakdown - Mailboxes: $([Math]::Round($Summary.MailboxStorageUsedMB / 1024, 2)) GB | OneDrive: $([Math]::Round($Summary.OneDriveStorageUsedMB / 1024, 2)) GB | Teams: $([Math]::Round($Summary.TeamsStorageUsedMB / 1024, 2)) GB" "INFO"
    }
    catch {
        Write-Log "Error creating consolidated report: $_" "ERROR"
    }
}

function New-ExecutionSummary {
    <#
    .SYNOPSIS
    Creates a summary report of the script execution.
    #>
    $EndTime = Get-Date
    $Duration = $EndTime - $Script:StartTime
    
    $Summary = @"
=== MICROSOFT 365 STORAGE REPORT EXECUTION SUMMARY ===
Start Time: $($Script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))
End Time: $($EndTime.ToString('yyyy-MM-dd HH:mm:ss'))
Duration: $($Duration.ToString('hh\:mm\:ss'))
Output Folder: $OutputFolder
Services Included: $(if($IncludeMailboxes){"Mailboxes "})$(if($IncludeOneDrive){"OneDrive "})$(if($IncludeTeams){"Teams"})
PowerShell Version: $($PSVersionTable.PSVersion)
=== END SUMMARY ===
"@
    
    Write-Log $Summary "INFO"
    
    $SummaryPath = Join-Path $OutputFolder "ExecutionSummary.txt"
    $Summary | Out-File -FilePath $SummaryPath -Encoding UTF8
    
    Write-Log "Execution summary saved to: $SummaryPath" "INFO"
}

#endregion

#region Main Execution

# Initialize output folder and logging
try {
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
        Write-Host "Created output folder: $OutputFolder" -ForegroundColor Green
    }
    
    $Script:LogFile = Join-Path $OutputFolder "M365StorageReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Write-Log "=== Microsoft 365 Storage Consumption Report Started ===" "INFO"
    Write-Log "Script version: 2.0" "INFO"
    Write-Log "Output folder: $OutputFolder" "INFO"
    Write-Log "Log file: $Script:LogFile" "INFO"
}
catch {
    Write-Error "Failed to initialize output folder: $_"
    exit 1
}

#region Pre-flight Checks and Initialization

# Pre-flight checks
try {
    Write-Log "Performing pre-flight checks..." "INFO"
    
    # 1. Verify PowerShell version
    Test-PowerShellVersion
    
    # 2. Check administrator privileges
    $IsAdmin = Test-AdminPrivileges
    
    # 3. Validate service selections
    if (-not $IncludeMailboxes -and -not $IncludeOneDrive -and -not $IncludeTeams) {
        throw "At least one service must be selected (Mailboxes, OneDrive, or Teams)."
    }
    
    Write-Log "Services to process: $(if($IncludeMailboxes){"Mailboxes "})$(if($IncludeOneDrive){"OneDrive "})$(if($IncludeTeams){"Teams"})" "INFO"
    
    # 4. Initialize required modules
    Initialize-RequiredModules -IsAdmin $IsAdmin
    
    Write-Log "Pre-flight checks completed successfully." "SUCCESS"
}
catch {
    Write-Log "Pre-flight checks failed: $_" "ERROR"
    exit 1
}

#endregion

# Collect data from each service
$AllReports = @()
$SuccessCount = 0
$ServiceCount = 0

if ($IncludeMailboxes) { $ServiceCount++ }
if ($IncludeOneDrive) { $ServiceCount++ }
if ($IncludeTeams) { $ServiceCount++ }

Write-Log "Starting data collection from $ServiceCount service(s)..." "INFO"

if ($IncludeMailboxes) {
    Write-Log "Processing Exchange Online mailboxes..." "INFO"
    try {
        $MailboxData = Export-MailboxStorageData
        if ($MailboxData -and $MailboxData.Count -gt 0) { 
            $AllReports += ,$MailboxData
            $SuccessCount++
            Write-Log "Mailbox data collection: SUCCESS" "SUCCESS"
        } else {
            Write-Log "Mailbox data collection: NO DATA" "WARN"
        }
    }
    catch {
        Write-Log "Mailbox data collection: FAILED - $_" "ERROR"
    }
}

if ($IncludeOneDrive) {
    Write-Log "Processing OneDrive accounts..." "INFO"
    try {
        $OneDriveData = Export-OneDriveStorageData
        if ($OneDriveData -and $OneDriveData.Count -gt 0) { 
            $AllReports += ,$OneDriveData
            $SuccessCount++
            Write-Log "OneDrive data collection: SUCCESS" "SUCCESS"
        } else {
            Write-Log "OneDrive data collection: NO DATA" "WARN"
        }
    }
    catch {
        Write-Log "OneDrive data collection: FAILED - $_" "ERROR"
    }
}

if ($IncludeTeams) {
    Write-Log "Processing Teams sites..." "INFO"
    try {
        $TeamsData = Export-TeamsStorageData
        if ($TeamsData -and $TeamsData.Count -gt 0) { 
            $AllReports += ,$TeamsData
            $SuccessCount++
            Write-Log "Teams data collection: SUCCESS" "SUCCESS"
        } else {
            Write-Log "Teams data collection: NO DATA" "WARN"
        }
    }
    catch {
        Write-Log "Teams data collection: FAILED - $_" "ERROR"
    }
}

# Create consolidated report
Write-Log "Data collection completed: $SuccessCount/$ServiceCount services processed successfully." "INFO"

if ($AllReports.Count -gt 0) {
    Create-ConsolidatedReport -AllData $AllReports
    Write-Log "Storage consumption report completed successfully!" "SUCCESS"
    $ExitCode = 0
}
else {
    Write-Log "No data collected from any service. Please check your permissions and service availability." "ERROR"
    Write-Log "Ensure you have the required permissions:" "INFO"
    Write-Log "- Exchange Online: Global Admin or Exchange Admin role" "INFO"
    Write-Log "- Microsoft Graph: User.Read.All, Reports.Read.All, Sites.Read.All" "INFO"
    $ExitCode = 1
}

# Generate execution summary
New-ExecutionSummary

# Cleanup connections
Write-Log "Performing cleanup..." "INFO"
try {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Write-Log "Disconnected from all services." "SUCCESS"
}
catch {
    Write-Log "Error during cleanup (non-critical): $_" "WARN"
}

Write-Log "=== Microsoft 365 Storage Report execution completed ===" "INFO"
Write-Log "Check the output folder for all generated reports: $OutputFolder" "INFO"

exit $ExitCode

#endregion
