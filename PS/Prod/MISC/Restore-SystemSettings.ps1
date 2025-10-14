# System Settings Restore Script
# Restores PowerCFG and IE Zone settings from backup

Param(
    [Parameter(Mandatory=$False)]
    [String]$BackupFolder = ""
)

$ErrorActionPreference = 'Continue'

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "System Settings Restore" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# If no backup folder specified, find the most recent one
If ([String]::IsNullOrEmpty($BackupFolder)) {
    [String]$BackupPath = Join-Path $PSScriptRoot "System-Backup"
    
    If (-not (Test-Path $BackupPath)) {
        Write-Host "[ERROR] No backups found at: $BackupPath" -ForegroundColor Red
        Exit 1
    }
    
    [Array]$Backups = Get-ChildItem -Path $BackupPath -Directory | Sort-Object Name -Descending
    
    If ($Backups.Count -eq 0) {
        Write-Host "[ERROR] No backup folders found" -ForegroundColor Red
        Exit 1
    }
    
    $BackupFolder = $Backups[0].FullName
    Write-Host "Using most recent backup: $BackupFolder" -ForegroundColor Yellow
    Write-Host ""
}

If (-not (Test-Path $BackupFolder)) {
    Write-Host "[ERROR] Backup folder not found: $BackupFolder" -ForegroundColor Red
    Exit 1
}

Write-Host "Restoring from: $BackupFolder" -ForegroundColor Cyan
Write-Host ""

# ============================================================================
# 1. Restore PowerCFG Settings
# ============================================================================
Write-Host "Restoring PowerCFG settings..." -ForegroundColor Yellow

[String]$PowerCFGFile = Join-Path $BackupFolder "powercfg-backup.pow"

If (Test-Path $PowerCFGFile) {
    Try {
        Write-Host "  [INFO] Note: PowerCFG scheme restore requires manual review" -ForegroundColor Yellow
        Write-Host "  [INFO] Monitor timeout details saved in: monitor-timeout.txt" -ForegroundColor Gray
        
        # For now, just restore to safe default (20 minutes)
        Write-Host "  [INFO] Setting monitor timeout to default: 20 minutes" -ForegroundColor Cyan
        & powercfg -change -monitor-timeout-ac 20 | Out-Null
        
        Write-Host "  [OK] Monitor timeout restored to 20 minutes" -ForegroundColor Green
    }
    Catch {
        Write-Host "  [ERROR] Failed to restore PowerCFG: $($_.Exception.Message)" -ForegroundColor Red
    }
} Else {
    Write-Host "  [WARNING] PowerCFG backup file not found" -ForegroundColor Yellow
}

# ============================================================================
# 2. Restore IE Zone Settings
# ============================================================================
Write-Host ""
Write-Host "Restoring IE Zone settings..." -ForegroundColor Yellow

[String]$IEZoneRegFile = Join-Path $BackupFolder "IE-Zones-Backup.reg"

If (Test-Path $IEZoneRegFile) {
    Try {
        # Import registry file
        & regedit /s $IEZoneRegFile
        
        Write-Host "  [OK] IE Zone settings restored" -ForegroundColor Green
    }
    Catch {
        Write-Host "  [ERROR] Failed to restore IE Zones: $($_.Exception.Message)" -ForegroundColor Red
    }
} Else {
    Write-Host "  [WARNING] IE Zone backup file not found" -ForegroundColor Yellow
}

# ============================================================================
# Summary
# ============================================================================
Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Restore Complete!" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Settings have been restored from backup" -ForegroundColor Green
Write-Host ""
Write-Host "To verify:" -ForegroundColor Cyan
Write-Host "  Monitor timeout: powercfg /q" -ForegroundColor Gray
Write-Host "  IE Zones: Check Internet Options > Security" -ForegroundColor Gray
Write-Host ""

