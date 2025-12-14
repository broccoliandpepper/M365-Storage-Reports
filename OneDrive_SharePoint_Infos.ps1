function Connect-ToMicrosoftGraph {
    <#
    .SYNOPSIS
    Connects to Microsoft Graph for Cloud Services Report with required permissions
    #>
    
    $ErrorActionPreference = 'Stop'
    
    Write-Host "=== Starting Microsoft Graph Connection ===" -ForegroundColor Cyan
    Write-Host "Current time: $(Get-Date)" -ForegroundColor Cyan

    # Check if Microsoft.Graph module is installed
    Write-Host "Step 1: Checking if Microsoft.Graph module is installed..." -ForegroundColor Yellow
    try {
        $module = Get-Module -ListAvailable -Name Microsoft.Graph -ErrorAction SilentlyContinue
        if (-not $module) {
            Write-Host "Microsoft.Graph module not found. Installing..." -ForegroundColor Yellow
            Write-Host "This may take a few minutes..." -ForegroundColor Yellow
            Install-Module Microsoft.Graph -Scope CurrentUser -Force -ErrorAction Stop
            Write-Host "Module installation completed." -ForegroundColor Green
        } else {
            Write-Host "Microsoft.Graph module found (Version: $($module[0].Version))" -ForegroundColor Green
        }
    } catch {
        Write-Host "ERROR during module check/install: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    # Import the module
    Write-Host "Step 2: Importing Microsoft.Graph module..." -ForegroundColor Yellow
    try {
        Import-Module Microsoft.Graph -ErrorAction Stop
        Write-Host "Module imported successfully." -ForegroundColor Green
    } catch {
        Write-Host "ERROR importing module: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    # Check if already connected with required scopes
    Write-Host "Step 3: Checking existing connection..." -ForegroundColor Yellow
    try {
        $context = Get-MgContext
        if ($context) {
            $requiredScopes = @("Sites.Read.All", "Team.ReadBasic.All", "Directory.Read.All", "Reports.Read.All", "AppCatalog.Read.All")
            $hasAllScopes = $true
            
            foreach ($scope in $requiredScopes) {
                if ($context.Scopes -notcontains $scope) {
                    Write-Host "Missing required scope: $scope" -ForegroundColor Yellow
                    $hasAllScopes = $false
                }
            }
            
            if ($hasAllScopes) {
                Write-Host "Already connected with all required permissions to tenant: $($context.TenantId)" -ForegroundColor Green
                Write-Host "Account: $($context.Account)" -ForegroundColor Green
                return $true
            } else {
                Write-Host "Connected but missing required permissions. Need to reconnect..." -ForegroundColor Yellow
                Disconnect-MgGraph -ErrorAction SilentlyContinue
            }
        } else {
            Write-Host "No existing connection found." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Not connected to any tenant. Error: $($_.Exception.Message)" -ForegroundColor Red
    }

    # Prompt for tenant ID
    Write-Host "Step 4: Requesting tenant ID..." -ForegroundColor Yellow
    try {
        $tenantId = Read-Host "Enter your Tenant ID"
        Write-Host "Received Tenant ID: $tenantId" -ForegroundColor Cyan
        
        if (-not ($tenantId -match '^[0-9a-fA-F-]{36}$')) {
            Write-Host "Invalid Tenant ID format. Expected format: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ForegroundColor Red
            return $false
        }
        Write-Host "Tenant ID format validation passed." -ForegroundColor Green
    } catch {
        Write-Host "ERROR getting tenant ID: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }

    # Connect to Graph with all required permissions
    Write-Host "Step 5: Connecting to Microsoft Graph with cloud services permissions..." -ForegroundColor Yellow
    try {
        $scopes = @(
            "Sites.Read.All",
            "Team.ReadBasic.All", 
            "Directory.Read.All",
            "Reports.Read.All",
            "AppCatalog.Read.All",
            "User.Read.All"
        )
        Write-Host "Connecting with scopes: $($scopes -join ', ')" -ForegroundColor Cyan
        Connect-MgGraph -TenantId $tenantId -Scopes $scopes -NoWelcome
        
        Write-Host "Connection attempt completed. Verifying..." -ForegroundColor Yellow
        $context = Get-MgContext
        if ($context) {
            Write-Host "Successfully connected to tenant: $($context.TenantId)" -ForegroundColor Green
            Write-Host "Account: $($context.Account)" -ForegroundColor Green
            Write-Host "Available scopes: $($context.Scopes -join ', ')" -ForegroundColor Green
            return $true
        } else {
            Write-Host "Connection failed - no context available." -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "Failed to connect. Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please check your credentials, tenant ID, and ensure you have admin permissions." -ForegroundColor Red
        return $false
    }
}

function Get-CloudServicesReport {
    <#
    .SYNOPSIS
    Collects information from OneDrive, SharePoint, and Teams:
    - Active OneDrive users and average size
    - SharePoint site names
    - Teams names and private channels
    - Third-party apps connected to Teams
    .OUTPUTS
    Table on screen and CSV file: ./Cloud_Infos.csv
    #>

    Write-Host "=== Starting Cloud Services Report ===" -ForegroundColor Cyan
    Write-Host "Current time: $(Get-Date)" -ForegroundColor Cyan

    # Check if we have the necessary Graph commands available
    Write-Host "Step 1: Checking available Graph commands..." -ForegroundColor Yellow
    $requiredCommands = @("Get-MgUser", "Get-MgSite", "Get-MgTeam", "Get-MgReportOneDriveUsageAccountDetail")
    $availableCommands = @()
    
    foreach ($command in $requiredCommands) {
        if (Get-Command $command -ErrorAction SilentlyContinue) {
            $availableCommands += $command
            Write-Host "Command $command is available." -ForegroundColor Green
        } else {
            Write-Host "Command $command is NOT available." -ForegroundColor Yellow
        }
    }

    $cloudData = @()

    # --- OneDrive Users and Size ---
    Write-Host "`nStep 2: Fetching OneDrive usage reports..." -ForegroundColor Yellow
    try {
        Write-Host "Retrieving OneDrive usage data for last 7 days..." -ForegroundColor Cyan
        $oneDriveReport = Get-MgReportOneDriveUsageAccountDetail -Period D7 -ErrorAction Stop
        
        if ($oneDriveReport) {
            $oneDriveCount = 0
            foreach ($entry in $oneDriveReport) {
                if ($entry.StorageUsedInBytes -gt 0) {  # Only active users
                    $oneDriveCount++
                    $sizeMB = [math]::Round($entry.StorageUsedInBytes / 1MB, 2)
                    $cloudData += [PSCustomObject]@{
                        Service     = "OneDrive"
                        Name        = $entry.UserPrincipalName
                        Detail      = "Active User"
                        SizeMB      = $sizeMB
                        LastActivity = $entry.LastActivityDate
                        URL         = $entry.SiteUrl
                    }
                }
            }
            Write-Host "Successfully retrieved $oneDriveCount active OneDrive users." -ForegroundColor Green
        } else {
            Write-Host "No OneDrive usage data returned." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "ERROR fetching OneDrive usage data: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "This might be due to insufficient permissions or reporting limitations." -ForegroundColor Yellow
    }

    # --- SharePoint Sites ---
    Write-Host "`nStep 3: Fetching SharePoint sites..." -ForegroundColor Yellow
    try {
        Write-Host "Retrieving SharePoint sites (limited to first 100)..." -ForegroundColor Cyan
        $sites = Get-MgSite -Search "*" -Top 100 -ErrorAction Stop
        
        $siteCount = 0
        foreach ($site in $sites) {
            $siteCount++
            if ($siteCount % 10 -eq 0) {
                Write-Host "Processed $siteCount SharePoint sites..." -ForegroundColor Cyan
            }
            
            $cloudData += [PSCustomObject]@{
                Service     = "SharePoint"
                Name        = $site.DisplayName
                Detail      = "Site"
                SizeMB      = "N/A"
                LastActivity = "N/A"
                URL         = $site.WebUrl
            }
        }
        Write-Host "Successfully retrieved $siteCount SharePoint sites." -ForegroundColor Green
    } catch {
        Write-Host "ERROR fetching SharePoint sites: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "This might be due to insufficient permissions (Sites.Read.All required)." -ForegroundColor Yellow
    }

    # --- Teams and Private Channels ---
    Write-Host "`nStep 4: Fetching Teams and channels..." -ForegroundColor Yellow
    try {
        Write-Host "Retrieving all Teams..." -ForegroundColor Cyan
        $teams = Get-MgTeam -All -ErrorAction Stop
        
        $teamCount = 0
        $channelCount = 0
        
        foreach ($team in $teams) {
            $teamCount++
            Write-Host "Processing team $teamCount of $($teams.Count): $($team.DisplayName)" -ForegroundColor Cyan
            
            # Add team info
            $cloudData += [PSCustomObject]@{
                Service     = "Teams"
                Name        = $team.DisplayName
                Detail      = "Team"
                SizeMB      = "N/A"
                LastActivity = "N/A"
                URL         = "N/A"
            }

            # Get channels for this team
            try {
                $channels = Get-MgTeamChannel -TeamId $team.Id -ErrorAction Stop
                foreach ($channel in $channels) {
                    if ($channel.MembershipType -eq "private") {
                        $channelCount++
                        $cloudData += [PSCustomObject]@{
                            Service     = "Teams"
                            Name        = "$($team.DisplayName) - $($channel.DisplayName)"
                            Detail      = "Private Channel"
                            SizeMB      = "N/A"
                            LastActivity = "N/A"
                            URL         = "N/A"
                        }
                    }
                }
            } catch {
                Write-Host "WARNING: Could not retrieve channels for team: $($team.DisplayName)" -ForegroundColor Yellow
            }
        }
        Write-Host "Successfully retrieved $teamCount Teams and $channelCount private channels." -ForegroundColor Green
    } catch {
        Write-Host "ERROR fetching Teams: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "This might be due to insufficient permissions (Team.ReadBasic.All required)." -ForegroundColor Yellow
    }

    # --- Third-party Apps in Teams ---
    Write-Host "`nStep 5: Fetching Teams applications..." -ForegroundColor Yellow
    try {
        Write-Host "Retrieving Teams app catalog..." -ForegroundColor Cyan
        $apps = Get-MgAppCatalogTeamApp -All -ErrorAction Stop
        
        $appCount = 0
        foreach ($app in $apps) {
            # Focus on third-party apps (not Microsoft built-in apps)
            if ($app.DistributionMethod -eq "store" -and $app.ExternalId -notlike "*microsoft*") {
                $appCount++
                $cloudData += [PSCustomObject]@{
                    Service     = "Teams"
                    Name        = $app.DisplayName
                    Detail      = "3rd Party App"
                    SizeMB      = "N/A"
                    LastActivity = "N/A"
                    URL         = "N/A"
                }
            }
        }
        Write-Host "Successfully retrieved $appCount third-party Teams apps." -ForegroundColor Green
    } catch {
        Write-Host "ERROR fetching Teams apps: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "This might be due to insufficient permissions (AppCatalog.Read.All required)." -ForegroundColor Yellow
    }

    # Display results
    Write-Host "`nStep 6: Displaying results..." -ForegroundColor Yellow
    if ($cloudData.Count -gt 0) {
        $cloudData | Format-Table -AutoSize
        
        # Export to CSV
        Write-Host "Step 7: Exporting to CSV..." -ForegroundColor Yellow
        try {
            $csvPath = "./Cloud_Infos.csv"
            $cloudData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
            Write-Host "Cloud services info exported to $csvPath" -ForegroundColor Green
            
            # Summary statistics
            $oneDriveUsers = ($cloudData | Where-Object { $_.Service -eq "OneDrive" }).Count
            $sharePointSites = ($cloudData | Where-Object { $_.Service -eq "SharePoint" }).Count
            $teamsCount = ($cloudData | Where-Object { $_.Service -eq "Teams" -and $_.Detail -eq "Team" }).Count
            $privateChannels = ($cloudData | Where-Object { $_.Detail -eq "Private Channel" }).Count
            $thirdPartyApps = ($cloudData | Where-Object { $_.Detail -eq "3rd Party App" }).Count
            $totalOneDriveSizeMB = ($cloudData | Where-Object { $_.Service -eq "OneDrive" -and $_.SizeMB -ne "N/A" } | Measure-Object -Property SizeMB -Sum).Sum
            
            Write-Host "`n=== SUMMARY ===" -ForegroundColor Magenta
            Write-Host "Active OneDrive Users: $oneDriveUsers" -ForegroundColor Cyan
            Write-Host "Total OneDrive Storage: $totalOneDriveSizeMB MB" -ForegroundColor Cyan
            Write-Host "SharePoint Sites: $sharePointSites" -ForegroundColor Cyan
            Write-Host "Teams: $teamsCount" -ForegroundColor Cyan
            Write-Host "Private Channels: $privateChannels" -ForegroundColor Cyan
            Write-Host "Third-party Apps: $thirdPartyApps" -ForegroundColor Cyan
            Write-Host "Total Items: $($cloudData.Count)" -ForegroundColor Cyan
            
            return $true
        } catch {
            Write-Host "ERROR exporting to CSV: $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    } else {
        Write-Host "No cloud services data retrieved. Check permissions and try again." -ForegroundColor Yellow
        return $false
    }
}

function Get-CloudServicesReport {
    <#
    .SYNOPSIS
    Main function that orchestrates the cloud services data collection
    #>
    
    Write-Host "=== Starting Cloud Services Data Collection ===" -ForegroundColor Cyan
    
    # First ensure we're connected with proper permissions
    $connectionResult = Connect-ToMicrosoftGraph
    
    if ($connectionResult -eq $true) {
        Write-Host "`n=== Connection successful. Proceeding with data collection ===" -ForegroundColor Green
        
        # Check if we have the necessary Graph commands available
        Write-Host "Step 1: Checking available Graph commands..." -ForegroundColor Yellow
        $requiredCommands = @(
            "Get-MgReportOneDriveUsageAccountDetail",
            "Get-MgSite", 
            "Get-MgTeam", 
            "Get-MgTeamChannel",
            "Get-MgAppCatalogTeamApp"
        )
        
        $availableCommands = @()
        foreach ($command in $requiredCommands) {
            if (Get-Command $command -ErrorAction SilentlyContinue) {
                $availableCommands += $command
                Write-Host "✓ Command $command is available." -ForegroundColor Green
            } else {
                Write-Host "✗ Command $command is NOT available." -ForegroundColor Yellow
            }
        }
        
        if ($availableCommands.Count -eq 0) {
            Write-Host "ERROR: No required Graph commands are available. Please check module installation." -ForegroundColor Red
            return $false
        }
        
        $cloudData = @()

        # --- OneDrive Users and Size ---
        Write-Host "`nStep 2: Fetching OneDrive usage reports..." -ForegroundColor Yellow
        if ("Get-MgReportOneDriveUsageAccountDetail" -in $availableCommands) {
            try {
                Write-Host "Retrieving OneDrive usage data for last 7 days..." -ForegroundColor Cyan
                $oneDriveReport = Get-MgReportOneDriveUsageAccountDetail -Period D7 -ErrorAction Stop
                
                if ($oneDriveReport) {
                    $oneDriveCount = 0
                    foreach ($entry in $oneDriveReport) {
                        if ($entry.StorageUsedInBytes -gt 0) {  # Only active users
                            $oneDriveCount++
                            $sizeMB = [math]::Round($entry.StorageUsedInBytes / 1MB, 2)
                            $cloudData += [PSCustomObject]@{
                                Service      = "OneDrive"
                                Name         = $entry.UserPrincipalName
                                Detail       = "Active User"
                                SizeMB       = $sizeMB
                                LastActivity = $entry.LastActivityDate
                                URL          = $entry.SiteUrl
                            }
                        }
                    }
                    Write-Host "Successfully retrieved $oneDriveCount active OneDrive users." -ForegroundColor Green
                } else {
                    Write-Host "No OneDrive usage data returned." -ForegroundColor Yellow
                }
            } catch {
                Write-Host "ERROR fetching OneDrive usage data: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "OneDrive reporting command not available. Skipping..." -ForegroundColor Yellow
        }

        # --- SharePoint Sites ---
        Write-Host "`nStep 3: Fetching SharePoint sites..." -ForegroundColor Yellow
        if ("Get-MgSite" -in $availableCommands) {
            try {
                Write-Host "Retrieving SharePoint sites (limited to first 100)..." -ForegroundColor Cyan
                $sites = Get-MgSite -Search "*" -Top 100 -ErrorAction Stop
                
                $siteCount = 0
                foreach ($site in $sites) {
                    $siteCount++
                    if ($siteCount % 10 -eq 0) {
                        Write-Host "Processed $siteCount SharePoint sites..." -ForegroundColor Cyan
                    }
                    
                    $cloudData += [PSCustomObject]@{
                        Service      = "SharePoint"
                        Name         = $site.DisplayName
                        Detail       = "Site"
                        SizeMB       = "N/A"
                        LastActivity = if ($site.LastModifiedDateTime) { $site.LastModifiedDateTime.ToString("yyyy-MM-dd") } else { "N/A" }
                        URL          = $site.WebUrl
                    }
                }
                Write-Host "Successfully retrieved $siteCount SharePoint sites." -ForegroundColor Green
            } catch {
                Write-Host "ERROR fetching SharePoint sites: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "SharePoint sites command not available. Skipping..." -ForegroundColor Yellow
        }

        # --- Teams and Private Channels ---
        Write-Host "`nStep 4: Fetching Teams and channels..." -ForegroundColor Yellow
        if ("Get-MgTeam" -in $availableCommands) {
            try {
                Write-Host "Retrieving all Teams..." -ForegroundColor Cyan
                $teams = Get-MgTeam -All -ErrorAction Stop
                
                $teamCount = 0
                $channelCount = 0
                
                foreach ($team in $teams) {
                    $teamCount++
                    if ($teamCount % 5 -eq 0) {
                        Write-Host "Processing team $teamCount of $($teams.Count)..." -ForegroundColor Cyan
                    }
                    
                    # Add team info
                    $cloudData += [PSCustomObject]@{
                        Service      = "Teams"
                        Name         = $team.DisplayName
                        Detail       = "Team"
                        SizeMB       = "N/A"
                        LastActivity = "N/A"
                        URL          = "N/A"
                    }

                    # Get channels for this team
                    if ("Get-MgTeamChannel" -in $availableCommands) {
                        try {
                            $channels = Get-MgTeamChannel -TeamId $team.Id -ErrorAction Stop
                            foreach ($channel in $channels) {
                                if ($channel.MembershipType -eq "private") {
                                    $channelCount++
                                    $cloudData += [PSCustomObject]@{
                                        Service      = "Teams"
                                        Name         = "$($team.DisplayName) - $($channel.DisplayName)"
                                        Detail       = "Private Channel"
                                        SizeMB       = "N/A"
                                        LastActivity = "N/A"
                                        URL          = "N/A"
                                    }
                                }
                            }
                        } catch {
                            Write-Host "WARNING: Could not retrieve channels for team: $($team.DisplayName)" -ForegroundColor Yellow
                        }
                    }
                }
                Write-Host "Successfully retrieved $teamCount Teams and $channelCount private channels." -ForegroundColor Green
            } catch {
                Write-Host "ERROR fetching Teams: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "Teams command not available. Skipping..." -ForegroundColor Yellow
        }

        # --- Third-party Apps in Teams ---
        Write-Host "`nStep 5: Fetching Teams applications..." -ForegroundColor Yellow
        if ("Get-MgAppCatalogTeamApp" -in $availableCommands) {
            try {
                Write-Host "Retrieving Teams app catalog..." -ForegroundColor Cyan
                $apps = Get-MgAppCatalogTeamApp -All -ErrorAction Stop
                
                $appCount = 0
                foreach ($app in $apps) {
                    # Focus on third-party apps (not Microsoft built-in apps)
                    if ($app.DistributionMethod -eq "store" -and $app.ExternalId -notlike "*microsoft*") {
                        $appCount++
                        $cloudData += [PSCustomObject]@{
                            Service      = "Teams"
                            Name         = $app.DisplayName
                            Detail       = "3rd Party App"
                            SizeMB       = "N/A"
                            LastActivity = "N/A"
                            URL          = "N/A"
                        }
                    }
                }
                Write-Host "Successfully retrieved $appCount third-party Teams apps." -ForegroundColor Green
            } catch {
                Write-Host "ERROR fetching Teams apps: $($_.Exception.Message)" -ForegroundColor Red
            }
        } else {
            Write-Host "Teams app catalog command not available. Skipping..." -ForegroundColor Yellow
        }

        # Display and export results
        Write-Host "`nStep 6: Processing results..." -ForegroundColor Yellow
        if ($cloudData.Count -gt 0) {
            Write-Host "Displaying $($cloudData.Count) cloud service items..." -ForegroundColor Cyan
            $cloudData | Format-Table -AutoSize

            # Export to CSV
            Write-Host "Step 7: Exporting to CSV..." -ForegroundColor Yellow
            try {
                $csvPath = "./Cloud_Infos.csv"
                $cloudData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                Write-Host "Cloud services info exported to $csvPath" -ForegroundColor Green
                
                # Summary statistics
                $oneDriveUsers = ($cloudData | Where-Object { $_.Service -eq "OneDrive" }).Count
                $sharePointSites = ($cloudData | Where-Object { $_.Service -eq "SharePoint" }).Count
                $teamsCount = ($cloudData | Where-Object { $_.Service -eq "Teams" -and $_.Detail -eq "Team" }).Count
                $privateChannels = ($cloudData | Where-Object { $_.Detail -eq "Private Channel" }).Count
                $thirdPartyApps = ($cloudData | Where-Object { $_.Detail -eq "3rd Party App" }).Count
                $totalOneDriveSizeMB = ($cloudData | Where-Object { $_.Service -eq "OneDrive" -and $_.SizeMB -ne "N/A" } | Measure-Object -Property SizeMB -Sum).Sum
                
                Write-Host "`n=== CLOUD SERVICES SUMMARY ===" -ForegroundColor Magenta
                Write-Host "Active OneDrive Users: $oneDriveUsers" -ForegroundColor Cyan
                Write-Host "Total OneDrive Storage: $totalOneDriveSizeMB MB" -ForegroundColor Cyan
                Write-Host "SharePoint Sites: $sharePointSites" -ForegroundColor Cyan
                Write-Host "Teams: $teamsCount" -ForegroundColor Cyan
                Write-Host "Private Channels: $privateChannels" -ForegroundColor Cyan
                Write-Host "Third-party Apps: $thirdPartyApps" -ForegroundColor Cyan
                Write-Host "Total Items Collected: $($cloudData.Count)" -ForegroundColor Cyan
                
                return $true
            } catch {
                Write-Host "ERROR exporting to CSV: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        } else {
            Write-Host "No cloud services data retrieved. This might be due to insufficient permissions." -ForegroundColor Yellow
            Write-Host "Required permissions: Sites.Read.All, Team.ReadBasic.All, Reports.Read.All, AppCatalog.Read.All" -ForegroundColor Yellow
            return $false
        }
    } else {
        Write-Host "Connection to Microsoft Graph failed. Cannot proceed with cloud services report." -ForegroundColor Red
        return $false
    }
}

# Execute the script with verbose output
Write-Host "=== Starting Cloud Services Report Script ===" -ForegroundColor Magenta
Write-Host "Current time: $(Get-Date)" -ForegroundColor Magenta

$result = Get-CloudServicesReport
Write-Host "`n=== Cloud Services Report completed. Result: $result ===" -ForegroundColor Magenta