# System Settings Backup Script
# Backs up PowerCFG and IE Zone settings before Phase 5 testing

$ErrorActionPreference = 'Continue'

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "System Settings Backup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

[String]$BackupPath = Join-Path $PSScriptRoot "System-Backup"
[String]$Timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
[String]$BackupFolder = Join-Path $BackupPath $Timestamp

# Create backup directory
If (-not (Test-Path $BackupFolder)) {
    New-Item -ItemType Directory -Path $BackupFolder -Force | Out-Null
}

Write-Host "Backup location: $BackupFolder" -ForegroundColor Cyan
Write-Host ""

# ============================================================================
# 1. Backup PowerCFG Settings
# ============================================================================
Write-Host "Backing up PowerCFG settings..." -ForegroundColor Yellow

Try {
    # Export current power scheme
    [String]$PowerCFGFile = Join-Path $BackupFolder "powercfg-backup.pow"
    
    # Get active power scheme GUID
    [String]$ActiveSchemeOutput = & powercfg /getactivescheme
    If ($ActiveSchemeOutput -match '([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})') {
        [String]$ActiveSchemeGUID = $Matches[1]
        
        # Export the scheme
        & powercfg /export $PowerCFGFile $ActiveSchemeGUID | Out-Null
        
        Write-Host "  [OK] PowerCFG scheme exported" -ForegroundColor Green
        Write-Host "       Active Scheme: $ActiveSchemeGUID" -ForegroundColor Gray
    } Else {
        Write-Host "  [WARNING] Could not determine active power scheme" -ForegroundColor Yellow
    }
    
    # Also save current monitor timeout settings
    [String]$MonitorTimeoutFile = Join-Path $BackupFolder "monitor-timeout.txt"
    & powercfg /q | Out-File -FilePath $MonitorTimeoutFile -Encoding UTF8
    
    Write-Host "  [OK] Monitor timeout settings saved" -ForegroundColor Green
}
Catch {
    Write-Host "  [ERROR] Failed to backup PowerCFG: $($_.Exception.Message)" -ForegroundColor Red
}

# ============================================================================
# 2. Backup IE Zone Settings
# ============================================================================
Write-Host ""
Write-Host "Backing up IE Zone settings..." -ForegroundColor Yellow

Try {
    [String]$IEZoneRegFile = Join-Path $BackupFolder "IE-Zones-Backup.reg"
    [String]$IEZoneRegPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones"
    
    # Export IE zone registry settings
    & reg export $IEZoneRegPath $IEZoneRegFile /y | Out-Null
    
    If (Test-Path $IEZoneRegFile) {
        Write-Host "  [OK] IE Zone settings exported" -ForegroundColor Green
        Write-Host "       Path: HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones" -ForegroundColor Gray
    } Else {
        Write-Host "  [WARNING] IE Zone registry export may have failed" -ForegroundColor Yellow
    }
}
Catch {
    Write-Host "  [ERROR] Failed to backup IE Zones: $($_.Exception.Message)" -ForegroundColor Red
}

# ============================================================================
# 3. Save Backup Metadata
# ============================================================================
[String]$MetadataFile = Join-Path $BackupFolder "backup-info.txt"

$BackupInfo = @"
Desktop Management Suite - System Settings Backup
Created: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Computer: $env:COMPUTERNAME
User: $env:USERNAME

Backed Up Settings:
- PowerCFG scheme and monitor timeout
- IE Zone configuration

To restore these settings:
.\Restore-SystemSettings.ps1 -BackupFolder "$BackupFolder"

Or manually restore:
- PowerCFG: Review $BackupFolder\monitor-timeout.txt
- IE Zones: regedit /s "$BackupFolder\IE-Zones-Backup.reg"
"@

$BackupInfo | Out-File -FilePath $MetadataFile -Encoding UTF8

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Backup Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Backup saved to:" -ForegroundColor Green
Write-Host "  $BackupFolder" -ForegroundColor Gray
Write-Host ""
Write-Host "Files created:" -ForegroundColor Cyan
Get-ChildItem -Path $BackupFolder | ForEach-Object {
    Write-Host "  - $($_.Name)" -ForegroundColor Gray
}
Write-Host ""
Write-Host "To restore after testing:" -ForegroundColor Yellow
Write-Host "  .\Restore-SystemSettings.ps1 -BackupFolder `"$BackupFolder`"" -ForegroundColor Gray
Write-Host ""

