#Export-MailboxStorage.ps1

#Requires -Version 7.3
<#
.SYNOPSIS
    Exports Exchange Online mailbox storage data to CSV.
.PARAMETER OutputFolder
    Path where the CSV report will be saved.
.EXAMPLE
    .\Export-MailboxStorage.ps1 -OutputFolder "C:\Reports"
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

function Install-ExchangeModule {
    try {
        $Module = Get-Module -ListAvailable -Name "ExchangeOnlineManagement" | Sort-Object Version -Descending | Select-Object -First 1
        
        if (-not $Module) {
            Write-Log "Installing ExchangeOnlineManagement module..." "INFO"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Install-Module -Name "ExchangeOnlineManagement" -Scope CurrentUser -Force -AllowClobber
            Write-Log "ExchangeOnlineManagement module installed successfully." "SUCCESS"
        } else {
            Write-Log "ExchangeOnlineManagement module (v$($Module.Version)) is available." "SUCCESS"
        }
    }
    catch {
        Write-Log "Failed to install ExchangeOnlineManagement module: $_" "ERROR"
        throw
    }
}

function Convert-BytesToMB {
    param($SizeValue)
    
    if ($SizeValue -eq $null -or $SizeValue -eq "" -or $SizeValue -eq "Unlimited") {
        return 0
    }
    
    if ($SizeValue -is [string]) {
        $BytesMatch = [regex]::Match($SizeValue, '\(([0-9,]+)\s+bytes\)')
        if ($BytesMatch.Success) {
            $BytesString = $BytesMatch.Groups[1].Value -replace ',', ''
            $Bytes = [long]$BytesString
            return [Math]::Round($Bytes / 1MB, 2)
        }
    }
    
    return 0
}

# Initialize
try {
    if (-not (Test-Path $OutputFolder)) {
        New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
    }
    
    $Script:LogFile = Join-Path $OutputFolder "MailboxStorage_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    Write-Log "=== Mailbox Storage Export Started ===" "INFO"
    
    # Install module
    Install-ExchangeModule
    
    # Connect to Exchange Online
    Write-Log "Connecting to Exchange Online..." "INFO"
    Import-Module ExchangeOnlineManagement -Force
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Log "Connected to Exchange Online." "SUCCESS"
    
    # Get mailboxes and statistics
    Write-Log "Retrieving mailbox list..." "INFO"
    $Mailboxes = Get-Mailbox -ResultSize Unlimited
    Write-Log "Found $($Mailboxes.Count) mailboxes to process." "INFO"
    
    $Report = @()
    $i = 0
    
    foreach ($Mailbox in $Mailboxes) {
        $i++
        Write-Progress -Activity "Processing Mailboxes" -Status "Processing $($Mailbox.DisplayName) ($i/$($Mailboxes.Count))" -PercentComplete (($i / $Mailboxes.Count) * 100)
        
        try {
            $Stats = Get-MailboxStatistics -Identity $Mailbox.DistinguishedName -WarningAction SilentlyContinue
            
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
            $Report += $ReportLine
        }
        catch {
            Write-Log "Error processing mailbox $($Mailbox.DisplayName): $_" "WARN"
        }
    }
    
    # Export results
    $OutputPath = Join-Path $OutputFolder "MailboxStorageData_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $Report | Sort-Object StorageUsedMB -Descending | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    
    Write-Log "Mailbox storage data exported to: $OutputPath" "SUCCESS"
    Write-Log "Total mailboxes processed: $($Report.Count)" "INFO"
    
    $TotalUsedGB = [Math]::Round(($Report | Measure-Object -Property StorageUsedMB -Sum).Sum / 1024, 2)
    Write-Log "Total mailbox storage used: $TotalUsedGB GB" "INFO"
    
    # Disconnect
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Log "=== Mailbox Storage Export Completed ===" "SUCCESS"
}
catch {
    Write-Log "Fatal error: $_" "ERROR"
    try { Disconnect-ExchangeOnline -Confirm:$false } catch { }
    exit 1
}
