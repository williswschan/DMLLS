<#
.SYNOPSIS
    Domain Features Validation Script
    
.DESCRIPTION
    Validates that all domain-specific features work correctly on a domain-joined computer.
    Run this script on a domain-joined computer to verify the PowerShell implementation.
    
.NOTES
    This script should ONLY be run on domain-joined computers!
    
.EXAMPLE
    .\Validate-DomainFeatures.ps1
#>

$ErrorActionPreference = 'Continue'

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Domain Features Validation" -ForegroundColor Cyan
Write-Host "Desktop Management Suite v2.0" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This script validates domain-specific functionality" -ForegroundColor Yellow
Write-Host "Run this on a DOMAIN-JOINED computer only!" -ForegroundColor Yellow
Write-Host ""

[Int]$ChecksPassed = 0
[Int]$ChecksFailed = 0
[Int]$ChecksWarning = 0

Function Write-CheckResult {
    Param(
        [String]$CheckName,
        [String]$Status,  # Pass, Fail, Warning
        [String]$Details = ""
    )
    
    Switch ($Status) {
        'Pass' {
            Write-Host "  [PASS] $CheckName" -ForegroundColor Green
            If ($Details) { Write-Host "         $Details" -ForegroundColor Gray }
            $Script:ChecksPassed++
        }
        'Fail' {
            Write-Host "  [FAIL] $CheckName" -ForegroundColor Red
            If ($Details) { Write-Host "         $Details" -ForegroundColor Red }
            $Script:ChecksFailed++
        }
        'Warning' {
            Write-Host "  [WARN] $CheckName" -ForegroundColor Yellow
            If ($Details) { Write-Host "         $Details" -ForegroundColor Yellow }
            $Script:ChecksWarning++
        }
    }
}

# Import modules
Import-Module .\Modules\Framework\DMLogger.psm1 -Force
Import-Module .\Modules\Utilities\Test-Environment.psm1 -Force
Import-Module .\Modules\Framework\DMComputer.psm1 -Force
Import-Module .\Modules\Framework\DMUser.psm1 -Force

# Initialize logging
Initialize-DMLog -JobType "DomainValidation" -VerboseLogging

# ============================================================================
# CHECK 1: Domain Join Status
# ============================================================================
Write-Host "=== Domain Join Status ===" -ForegroundColor Yellow

Try {
    [String]$Domain = $env:USERDNSDOMAIN
    
    If ([String]::IsNullOrEmpty($Domain)) {
        Write-CheckResult -CheckName "Domain Join" -Status "Fail" -Details "Computer is NOT domain-joined!"
        Write-Host ""
        Write-Host "This script must be run on a domain-joined computer!" -ForegroundColor Red
        Exit 1
    } Else {
        Write-CheckResult -CheckName "Domain Join" -Status "Pass" -Details "Domain: $Domain"
    }
} Catch {
    Write-CheckResult -CheckName "Domain Join" -Status "Fail" -Details $_.Exception.Message
}

# ============================================================================
# CHECK 2: Computer Information
# ============================================================================
Write-Host ""
Write-Host "=== Computer Information ===" -ForegroundColor Yellow

$Computer = Get-DMComputerInfo

# Check Distinguished Name
If ([String]::IsNullOrEmpty($Computer.DistinguishedName)) {
    Write-CheckResult -CheckName "Computer DN" -Status "Fail" -Details "DN is empty"
} Else {
    Write-CheckResult -CheckName "Computer DN" -Status "Pass" -Details $Computer.DistinguishedName
}

# Check Site
If ([String]::IsNullOrEmpty($Computer.Site)) {
    Write-CheckResult -CheckName "Computer Site" -Status "Warning" -Details "Site is empty"
} Else {
    Write-CheckResult -CheckName "Computer Site" -Status "Pass" -Details $Computer.Site
}

# Check City Code
If ($Computer.CityCode -eq "unknown") {
    Write-CheckResult -CheckName "Computer City Code" -Status "Warning" -Details "Could not extract city code from DN"
} Else {
    Write-CheckResult -CheckName "Computer City Code" -Status "Pass" -Details $Computer.CityCode
}

# Check Groups
If ($Null -eq $Computer.Groups -or $Computer.Groups.Count -eq 0) {
    Write-CheckResult -CheckName "Computer Groups" -Status "Warning" -Details "No groups found (may be normal)"
} Else {
    Write-CheckResult -CheckName "Computer Groups" -Status "Pass" -Details "$($Computer.Groups.Count) group(s) found"
    
    Write-Host "    First 5 groups:" -ForegroundColor Cyan
    For ([Int]$i = 0; $i -lt [Math]::Min(5, $Computer.Groups.Count); $i++) {
        Write-Host "      - $($Computer.Groups[$i].GroupName)" -ForegroundColor Gray
    }
}

# Check OU Mapping
If ([String]::IsNullOrEmpty($Computer.OUMapping)) {
    Write-CheckResult -CheckName "Computer OU Mapping" -Status "Warning" -Details "OU Mapping is empty"
} Else {
    Write-CheckResult -CheckName "Computer OU Mapping" -Status "Pass" -Details $Computer.OUMapping
}

# ============================================================================
# CHECK 3: User Information
# ============================================================================
Write-Host ""
Write-Host "=== User Information ===" -ForegroundColor Yellow

$User = Get-DMUserInfo

# Check Distinguished Name
If ([String]::IsNullOrEmpty($User.DistinguishedName)) {
    Write-CheckResult -CheckName "User DN" -Status "Fail" -Details "DN is empty"
} Else {
    Write-CheckResult -CheckName "User DN" -Status "Pass" -Details $User.DistinguishedName
}

# Check City Code
If ($User.CityCode -eq "unknown") {
    Write-CheckResult -CheckName "User City Code" -Status "Warning" -Details "Could not extract city code from DN"
} Else {
    Write-CheckResult -CheckName "User City Code" -Status "Pass" -Details $User.CityCode
}

# Check Groups
If ($Null -eq $User.Groups -or $User.Groups.Count -eq 0) {
    Write-CheckResult -CheckName "User Groups" -Status "Warning" -Details "No groups found (may be normal)"
} Else {
    Write-CheckResult -CheckName "User Groups" -Status "Pass" -Details "$($User.Groups.Count) group(s) found"
    
    Write-Host "    First 5 groups:" -ForegroundColor Cyan
    For ([Int]$i = 0; $i -lt [Math]::Min(5, $User.Groups.Count); $i++) {
        Write-Host "      - $($User.Groups[$i].GroupName)" -ForegroundColor Gray
    }
}

# Check OU Mapping
If ([String]::IsNullOrEmpty($User.OUMapping)) {
    Write-CheckResult -CheckName "User OU Mapping" -Status "Warning" -Details "OU Mapping is empty"
} Else {
    Write-CheckResult -CheckName "User OU Mapping" -Status "Pass" -Details $User.OUMapping
}

# ============================================================================
# CHECK 4: Environment Detection
# ============================================================================
Write-Host ""
Write-Host "=== Environment Detection ===" -ForegroundColor Yellow

# VPN Detection
[Boolean]$IsVPN = Test-DMVPNConnection
Write-CheckResult -CheckName "VPN Detection" -Status "Pass" -Details "VPN Connected: $IsVPN"

# Retail Detection
[Boolean]$IsRetailUser = Test-DMRetailUser -DistinguishedName $User.DistinguishedName
[Boolean]$IsRetailComp = Test-DMRetailComputer -DistinguishedName $Computer.DistinguishedName
Write-CheckResult -CheckName "Retail Detection" -Status "Pass" -Details "User: $IsRetailUser, Computer: $IsRetailComp"

# Terminal Session
[Object]$SessionInfo = Test-DMTerminalSession
Write-CheckResult -CheckName "Session Detection" -Status "Pass" -Details "Type: $($SessionInfo.SessionType), Terminal: $($SessionInfo.IsTerminalSession)"

# VM Detection
[Object]$VMInfo = Test-DMVirtualMachine
Write-CheckResult -CheckName "VM Detection" -Status "Pass" -Details "Platform: $($VMInfo.Platform), IsVirtual: $($VMInfo.IsVirtual)"

# ============================================================================
# CHECK 5: LDAP/AD Connectivity
# ============================================================================
Write-Host ""
Write-Host "=== LDAP/AD Connectivity ===" -ForegroundColor Yellow

# Test ADSystemInfo COM
Try {
    [Object]$ADSysInfo = New-Object -ComObject ADSystemInfo
    [String]$ForestName = $ADSysInfo.ForestDNSName
    Write-CheckResult -CheckName "ADSystemInfo COM" -Status "Pass" -Details "Forest: $ForestName"
} Catch {
    Write-CheckResult -CheckName "ADSystemInfo COM" -Status "Fail" -Details $_.Exception.Message
}

# Test ADODB Connection
Try {
    [Object]$Connection = New-Object -ComObject ADODB.Connection
    $Connection.Open("Provider=ADsDSOObject;")
    $Connection.Close()
    Write-CheckResult -CheckName "ADODB Connection" -Status "Pass" -Details "Can create LDAP connections"
} Catch {
    Write-CheckResult -CheckName "ADODB Connection" -Status "Fail" -Details $_.Exception.Message
}

# Test User Email Lookup
Try {
    [String]$Email = Get-DMUserEmail -DistinguishedName $User.DistinguishedName -Domain $User.Domain
    
    If ([String]::IsNullOrEmpty($Email)) {
        Write-CheckResult -CheckName "User Email (LDAP)" -Status "Warning" -Details "No email found (may be normal)"
    } Else {
        Write-CheckResult -CheckName "User Email (LDAP)" -Status "Pass" -Details $Email
    }
} Catch {
    Write-CheckResult -CheckName "User Email (LDAP)" -Status "Fail" -Details $_.Exception.Message
}

# Test Password Expiry
Try {
    [Object]$PwdInfo = Get-DMUserPasswordExpiry -DistinguishedName $User.DistinguishedName -Domain $User.Domain
    
    If ($PwdInfo.PasswordNeverExpires) {
        Write-CheckResult -CheckName "Password Expiry (LDAP)" -Status "Pass" -Details "Password never expires"
    } ElseIf ($PwdInfo.DaysUntilExpiry -lt 0) {
        Write-CheckResult -CheckName "Password Expiry (LDAP)" -Status "Warning" -Details "Could not calculate expiry"
    } Else {
        Write-CheckResult -CheckName "Password Expiry (LDAP)" -Status "Pass" -Details "$($PwdInfo.DaysUntilExpiry) days until expiry"
    }
} Catch {
    Write-CheckResult -CheckName "Password Expiry (LDAP)" -Status "Fail" -Details $_.Exception.Message
}

# ============================================================================
# CHECK 6: Backend Connectivity
# ============================================================================
Write-Host ""
Write-Host "=== Backend Service Connectivity ===" -ForegroundColor Yellow

Import-Module .\Modules\Services\DMServiceCommon.psm1 -Force

# Test Mapper Service
Try {
    [Object]$MapperServer = Get-DMServiceServer -ServiceName "ClassicMapper.asmx" -Domain $Computer.Domain
    
    If ($MapperServer.ServiceAvailable) {
        Write-CheckResult -CheckName "Mapper Service" -Status "Pass" -Details "$($MapperServer.ServerName) (Response: $($MapperServer.ResponseTime)ms)"
    } Else {
        Write-CheckResult -CheckName "Mapper Service" -Status "Warning" -Details "Service not available or not responding"
    }
} Catch {
    Write-CheckResult -CheckName "Mapper Service" -Status "Fail" -Details $_.Exception.Message
}

# Test Inventory Service
Try {
    [Object]$InventoryServer = Get-DMServiceServer -ServiceName "ClassicInventory.asmx" -Domain $Computer.Domain
    
    If ($InventoryServer.ServiceAvailable) {
        Write-CheckResult -CheckName "Inventory Service" -Status "Pass" -Details "$($InventoryServer.ServerName) (Response: $($InventoryServer.ResponseTime)ms)"
    } Else {
        Write-CheckResult -CheckName "Inventory Service" -Status "Warning" -Details "Service not available or not responding"
    }
} Catch {
    Write-CheckResult -CheckName "Inventory Service" -Status "Fail" -Details $_.Exception.Message
}

# ============================================================================
# Summary
# ============================================================================
Export-DMLog

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Validation Summary" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Passed:  $ChecksPassed" -ForegroundColor Green
Write-Host "Failed:  $ChecksFailed" -ForegroundColor $(If ($ChecksFailed -gt 0) { "Red" } Else { "Green" })
Write-Host "Warning: $ChecksWarning" -ForegroundColor Yellow
Write-Host ""

If ($ChecksFailed -eq 0) {
    Write-Host "[SUCCESS] All critical domain features validated!" -ForegroundColor Green
    Write-Host ""
    Write-Host "The PowerShell implementation should work correctly on this domain computer." -ForegroundColor Green
} Else {
    Write-Host "[WARNING] Some critical checks failed!" -ForegroundColor Red
    Write-Host ""
    Write-Host "There may be issues with domain connectivity or AD access." -ForegroundColor Yellow
    Write-Host "Review the failed checks above." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Detailed results saved to:" -ForegroundColor Cyan
Write-Host "  $(Get-DMLogPath)" -ForegroundColor Gray
Write-Host ""

# Save detailed report
$ReportPath = Join-Path $PSScriptRoot "Domain-Validation-Report.json"

$ValidationReport = [PSCustomObject]@{
    ValidationDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    ComputerName = $env:COMPUTERNAME
    UserName = $env:USERNAME
    Domain = $env:USERDNSDOMAIN
    ChecksPassed = $ChecksPassed
    ChecksFailed = $ChecksFailed
    ChecksWarning = $ChecksWarning
    ComputerInfo = $Computer
    UserInfo = $User
}

$ValidationReport | ConvertTo-Json -Depth 5 | Out-File -FilePath $ReportPath -Encoding UTF8

Write-Host "Validation report saved to:" -ForegroundColor Cyan
Write-Host "  $ReportPath" -ForegroundColor Gray
Write-Host ""

