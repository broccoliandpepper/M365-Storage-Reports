function Get-ExchangeMailboxReport {
    <#
    .SYNOPSIS
    Retrieves mailbox information from Exchange Online:
    - Mailbox name, average size, shared mailbox status
    - Outputs to table and CSV file

    .OUTPUTS
    Table on screen and CSV file: ./Exchange_Infos.csv
    #>

    $ErrorActionPreference = 'Stop'
    
    Write-Host "=== Starting Exchange Mailbox Report ===" -ForegroundColor Cyan
    Write-Host "Current time: $(Get-Date)" -ForegroundColor Cyan

    # Ensure ExchangeOnlineManagement module is available
    Write-Host "Step 1: Checking ExchangeOnlineManagement module..." -ForegroundColor Yellow
    try {
        $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue
        if (-not $module) {
            Write-Host "ExchangeOnlineManagement module not found. Installing..." -ForegroundColor Yellow
            Write-Host "This may take a few minutes..." -ForegroundColor Yellow
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "Module installation completed." -ForegroundColor Green
        } else {
            Write-Host "ExchangeOnlineManagement module found (Version: $($module[0].Version))" -ForegroundColor Green
        }
        
        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        Write-Host "Module imported successfully." -ForegroundColor Green
    } catch {
        Write-Host "ERROR during module check/install/import: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    # Connect to Exchange Online if not already connected
    Write-Host "Step 2: Checking Exchange Online connection..." -ForegroundColor Yellow
    try {
        # Test connection by trying to get a single mailbox
        $testConnection = Get-EXOMailbox -ResultSize 1 -ErrorAction Stop | Select-Object -First 1
        if ($testConnection) {
            Write-Host "Already connected to Exchange Online." -ForegroundColor Green
            $connectionInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($connectionInfo) {
                Write-Host "Connected as: $($connectionInfo.UserPrincipalName)" -ForegroundColor Green
                Write-Host "Tenant: $($connectionInfo.TenantId)" -ForegroundColor Green
            }
        }
    } catch {
        Write-Host "Not connected to Exchange Online. Connecting..." -ForegroundColor Yellow
        try {
            $adminUPN = Read-Host "Enter your admin UPN for Exchange Online connection"
            Write-Host "Connecting to Exchange Online as: $adminUPN" -ForegroundColor Cyan
            Connect-ExchangeOnline -UserPrincipalName $adminUPN -ErrorAction Stop
            Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
        } catch {
            Write-Host "ERROR connecting to Exchange Online: $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Please check your credentials and ensure you have Exchange Online admin rights." -ForegroundColor Red
            return $false
        }
    }

    # Get all mailboxes
    Write-Host "Step 3: Retrieving all mailboxes from Exchange Online..." -ForegroundColor Yellow
    try {
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop
        Write-Host "Successfully retrieved $($mailboxes.Count) mailboxes." -ForegroundColor Green
    } catch {
        Write-Host "ERROR retrieving mailboxes: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    # Prepare results
    Write-Host "Step 4: Processing mailbox data and statistics..." -ForegroundColor Yellow
    $results = @()
    $processedCount = 0
    $errorCount = 0
    
    foreach ($mb in $mailboxes) {
        $processedCount++
        if ($processedCount % 5 -eq 0) {
            Write-Host "Processed $processedCount of $($mailboxes.Count) mailboxes..." -ForegroundColor Cyan
        }
        
        try {
            # Get mailbox statistics
            $stats = Get-EXOMailboxStatistics -Identity $mb.UserPrincipalName -ErrorAction Stop
            
            # Calculate average size
            if ($stats.TotalItemSize -and $stats.TotalItemSize.Value) {
                try {
                    $avgSizeMB = [math]::Round($stats.TotalItemSize.Value.ToMB(), 2)
                } catch {
                    # Fallback calculation if ToMB() method fails
                    $sizeInBytes = $stats.TotalItemSize.Value.ToString() -replace '[^\d]', ''
                    if ($sizeInBytes -match '^\d+$') {
                        $avgSizeMB = [math]::Round([long]$sizeInBytes / 1MB, 2)
                    } else {
                        $avgSizeMB = 0
                    }
                }
            } else {
                $avgSizeMB = 0
            }
            
            # Check if shared mailbox
            $isShared = if ($mb.RecipientTypeDetails -eq "SharedMailbox") { "Yes" } else { "No" }
            
            # Get additional useful info
            $itemCount = if ($stats.ItemCount) { $stats.ItemCount } else { 0 }
            $lastLogon = if ($stats.LastLogonTime) { $stats.LastLogonTime.ToString("yyyy-MM-dd HH:mm") } else { "Never" }

            $results += [PSCustomObject]@{
                DisplayName       = $mb.DisplayName
                UserPrincipalName = $mb.UserPrincipalName
                AverageSizeMB     = $avgSizeMB
                ItemCount         = $itemCount
                SharedMailbox     = $isShared
                RecipientType     = $mb.RecipientTypeDetails
                LastLogon         = $lastLogon
                PrimaryEmail      = $mb.PrimarySmtpAddress
            }
        } catch {
            $errorCount++
            Write-Host "WARNING: Error processing mailbox $($mb.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Yellow
            
            # Add basic info even if stats retrieval failed
            $results += [PSCustomObject]@{
                DisplayName       = $mb.DisplayName
                UserPrincipalName = $mb.UserPrincipalName
                AverageSizeMB     = "Error"
                ItemCount         = "Error"
                SharedMailbox     = if ($mb.RecipientTypeDetails -eq "SharedMailbox") { "Yes" } else { "No" }
                RecipientType     = $mb.RecipientTypeDetails
                LastLogon         = "Error"
                PrimaryEmail      = $mb.PrimarySmtpAddress
            }
        }
    }

    Write-Host "Step 5: Displaying results..." -ForegroundColor Yellow
    # Output to screen
    $results | Format-Table -AutoSize

    # Export to CSV
    Write-Host "Step 6: Exporting to CSV..." -ForegroundColor Yellow
    try {
        $csvPath = "./Exchange_Infos.csv"
        $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exchange mailbox info exported to $csvPath" -ForegroundColor Green
        
        # Summary statistics
        $sharedCount = ($results | Where-Object { $_.SharedMailbox -eq "Yes" }).Count
        $userCount = ($results | Where-Object { $_.SharedMailbox -eq "No" }).Count
        $totalSizeMB = ($results | Where-Object { $_.AverageSizeMB -ne "Error" } | Measure-Object -Property AverageSizeMB -Sum).Sum
        $avgMailboxSize = if ($results.Count -gt 0) { [math]::Round($totalSizeMB / $results.Count, 2) } else { 0 }
        
        Write-Host "=== SUMMARY ===" -ForegroundColor Magenta
        Write-Host "Total Mailboxes: $($results.Count)" -ForegroundColor Cyan
        Write-Host "User Mailboxes: $userCount" -ForegroundColor Cyan
        Write-Host "Shared Mailboxes: $sharedCount" -ForegroundColor Cyan
        Write-Host "Total Storage Used: $totalSizeMB MB" -ForegroundColor Cyan
        Write-Host "Average Mailbox Size: $avgMailboxSize MB" -ForegroundColor Cyan
        if ($errorCount -gt 0) {
            Write-Host "Errors encountered: $errorCount" -ForegroundColor Yellow
        }
        
        return $true
    } catch {
        Write-Host "ERROR exporting to CSV: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Execute the function with verbose output
Write-Host "=== Starting Exchange Mailbox Report Script ===" -ForegroundColor Magenta
Write-Host "Current time: $(Get-Date)" -ForegroundColor Magenta

$result = Get-ExchangeMailboxReport
Write-Host "=== Exchange Mailbox Report completed. Result: $result ===" -ForegroundColor Magenta

# Optional: Disconnect from Exchange Online when done
Write-Host "`nWould you like to disconnect from Exchange Online? (Y/N): " -ForegroundColor Yellow -NoNewline
$disconnect = Read-Host
if ($disconnect -eq "Y" -or $disconnect -eq "y") {
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "Disconnected from Exchange Online." -ForegroundColor Green
    } catch {
        Write-Host "Note: Could not disconnect or already disconnected." -ForegroundColor Yellow
    }
}