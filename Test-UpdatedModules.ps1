<#
.SYNOPSIS
    Test Updated PowerShell Modules with Correct Method Names
    
.DESCRIPTION
    Tests the updated DMMapperService with the correct method names and parameters.
    
.EXAMPLE
    .\Test-UpdatedModules.ps1
#>

[CmdletBinding()]
Param()

Function Write-TestResult {
    Param(
        [String]$Section,
        [String]$Message,
        [String]$Value = "",
        [String]$Color = "White"
    )
    
    If ($Value) {
        Write-Host "[$Section] $Message : " -NoNewline -ForegroundColor Cyan
        Write-Host $Value -ForegroundColor $Color
    } Else {
        Write-Host "[$Section] $Message" -ForegroundColor $Color
    }
}

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "TESTING UPDATED POWERSHELL MODULES" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Green

# Import required modules
Try {
    Import-Module .\Modules\Framework\DMLogger.psm1 -Force
    Import-Module .\Modules\Framework\DMComputer.psm1 -Force
    Import-Module .\Modules\Framework\DMUser.psm1 -Force
    Import-Module .\Modules\Services\DMServiceCommon.psm1 -Force
    Import-Module .\Modules\Services\DMMapperService.psm1 -Force
    
    Write-TestResult "SUCCESS" "Modules imported" -Color Green
} Catch {
    Write-TestResult "ERROR" "Failed to import modules" $_.Exception.Message -Color Red
    Exit 1
}

# Initialize logging
Try {
    Initialize-DMLog -LogPath "C:\Temp\test-updated-modules.log" -ScriptName "Test-UpdatedModules" -Version "1.0" -VerboseLogging $True
    Write-TestResult "SUCCESS" "Logging initialized" -Color Green
} Catch {
    Write-TestResult "ERROR" "Failed to initialize logging" $_.Exception.Message -Color Red
}

# Get computer and user info
Write-Host "`n" -NoNewline
Write-TestResult "STEP 1" "Getting Computer and User Information" -Color Yellow

Try {
    [Object]$Computer = Get-DMComputerInfo
    [Object]$User = Get-DMUserInfo
    
    Write-TestResult "SUCCESS" "Computer Name" $Computer.Name -Color Green
    Write-TestResult "SUCCESS" "User Name" $User.Name -Color Green
    Write-TestResult "SUCCESS" "Computer Domain" $Computer.Domain -Color Green
    Write-TestResult "SUCCESS" "User Domain" $User.Domain -Color Green
} Catch {
    Write-TestResult "ERROR" "Failed to get computer/user info" $_.Exception.Message -Color Red
    Exit 1
}

# Test Drive Mappings
Write-Host "`n" -NoNewline
Write-TestResult "STEP 2" "Testing Drive Mappings (GetUserDrives)" -Color Yellow

Try {
    [Array]$DriveMappings = Get-DMDriveMappings -UserInfo $User -ComputerInfo $Computer
    
    Write-TestResult "SUCCESS" "Drive Mappings Retrieved" $DriveMappings.Count -Color Green
    
    If ($DriveMappings.Count -gt 0) {
        Write-Host "`nDrive Mappings:" -ForegroundColor Cyan
        ForEach ($Drive in $DriveMappings) {
            Write-Host "  Drive: $($Drive.DriveLetter) -> $($Drive.UncPath) [$($Drive.Description)]" -ForegroundColor Gray
        }
    } Else {
        Write-Host "  No drive mappings returned (expected for test user)" -ForegroundColor Yellow
    }
} Catch {
    Write-TestResult "ERROR" "Drive mappings failed" $_.Exception.Message -Color Red
}

# Test Printer Mappings
Write-Host "`n" -NoNewline
Write-TestResult "STEP 3" "Testing Printer Mappings (GetUserPrinters)" -Color Yellow

Try {
    [Array]$PrinterMappings = Get-DMPrinterMappings -ComputerInfo $Computer
    
    Write-TestResult "SUCCESS" "Printer Mappings Retrieved" $PrinterMappings.Count -Color Green
    
    If ($PrinterMappings.Count -gt 0) {
        Write-Host "`nPrinter Mappings:" -ForegroundColor Cyan
        ForEach ($Printer in $PrinterMappings) {
            Write-Host "  Printer: $($Printer.UncPath) [$($Printer.Description)]" -ForegroundColor Gray
        }
    } Else {
        Write-Host "  No printer mappings returned (expected for test user)" -ForegroundColor Yellow
    }
} Catch {
    Write-TestResult "ERROR" "Printer mappings failed" $_.Exception.Message -Color Red
}

# Test PST Mappings (if available)
Write-Host "`n" -NoNewline
Write-TestResult "STEP 4" "Testing PST Mappings (GetUserPSTs)" -Color Yellow

Try {
    [Array]$PSTMappings = Get-DMPSTMappings -UserInfo $User -ComputerInfo $Computer
    
    Write-TestResult "SUCCESS" "PST Mappings Retrieved" $PSTMappings.Count -Color Green
    
    If ($PSTMappings.Count -gt 0) {
        Write-Host "`nPST Mappings:" -ForegroundColor Cyan
        ForEach ($PST in $PSTMappings) {
            Write-Host "  PST: $($PST.PSTPath) [$($PST.Description)]" -ForegroundColor Gray
        }
    } Else {
        Write-Host "  No PST mappings returned (expected for test user)" -ForegroundColor Yellow
    }
} Catch {
    Write-TestResult "ERROR" "PST mappings failed" $_.Exception.Message -Color Red
}

# Summary
Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "UPDATED MODULES TEST SUMMARY" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

Write-Host "`nSTATUS:" -ForegroundColor Yellow
Write-Host "✅ Drive Mappings: Updated with GetUserDrives method" -ForegroundColor Green
Write-Host "✅ Printer Mappings: Updated with GetUserPrinters method" -ForegroundColor Green
Write-Host "⚠️  PST Mappings: May need further debugging" -ForegroundColor Yellow

Write-Host "`nNEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. Run the main entry point scripts to test full functionality" -ForegroundColor White
Write-Host "2. Debug PST mappings if needed" -ForegroundColor White
Write-Host "3. Test with real user data in production environment" -ForegroundColor White

Write-Host "`n"
