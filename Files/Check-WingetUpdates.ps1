<<<<<<< Updated upstream
# Check-WingetUpdates.ps1
=======
﻿# Check-WingetUpdates.ps1
>>>>>>> Stashed changes
<#
param(
    [Parameter(Mandatory = $true)][string]$Tenant,
    [Parameter(Mandatory = $true)][string]$AppId
)
#>

$AppId="05899b9b-7187-4a0c-a886-0506fb8c77ba"
$Tenant="daslab.onmicrosoft.com"

$global:AppId = $AppId

function Connect-Intune {
    param (
        [Parameter(Mandatory = $true)]
        [string]$tenant
    )

    Write-Host "Connecting to Intune..."

    if (-not $global:AppId) {
        Write-Warning "AppId not specified. Please execute 'Configure Intune integration'."
        return $false
    }

    Import-Module IntuneWin32App -ErrorAction SilentlyContinue

    if (Get-Module -Name IntuneWin32App) {
        try {
            Connect-MSIntuneGraph -TenantID $tenant -ClientId $global:AppId
            Write-Host "✅ Connected to Intune."
            return $true
        } catch {
            Write-Warning "❌ Failed to connect to Intune: $_"
            return $false
        }
    } else {
        Write-Warning "❌ Module IntuneWin32App not loaded."
        return $false
    }
}

function Get-LatestWingetVersion {
    param(
        [Parameter(Mandatory = $true)][string]$WingetId
    )

    try {
        $info = winget show --id "$WingetId" --accept-source-agreements 2>&1
        $versionLine = $info | Select-String -Pattern '^Version:' | Select-Object -First 1
        if ($versionLine) {
            return ($versionLine -replace 'Version:\s*', '').Trim()
        }
    } catch {
        Write-Warning "Failed to fetch winget info for $WingetId"
    }
    return $null
}

function Get-WingetManagedIntuneApps {
    $apps = Get-IntuneWin32App
    return $apps | Where-Object {
        $_.notes -match "WinGetID=(.*?)\s\|" -and $_.notes -match "Update=True"
    }
}

function Check-WingetUpdates {
    $apps = Get-WingetManagedIntuneApps

    foreach ($app in $apps) {
        if ($app.notes -match "WinGetID=(.*?)\s\|") {
            $wingetId = $matches[1].Trim()
            $currentVersion = $app.displayVersion

            Write-Host "🔍 Checking $($app.displayName) with WinGet ID: $wingetId (Current: $currentVersion)"

            $latestVersion = Get-LatestWingetVersion -WingetId $wingetId
            if ($latestVersion) {
                if ($latestVersion -ne $currentVersion) {
                    Write-Host "`n⚠️ Update available for $($app.displayName): $currentVersion -> $latestVersion`n" -ForegroundColor Yellow
                } else {
                    Write-Host "✅ Up-to-date: $($app.displayName) ($currentVersion)"
                }
            } else {
                Write-Host "❌ Failed to get version info for $wingetId"
            }
        }
    }
}

# Main execution
if (-not (Connect-Intune -tenant $Tenant)) {
    exit 1
}

Check-WingetUpdates

Write-Host "`n🎉 Done."
