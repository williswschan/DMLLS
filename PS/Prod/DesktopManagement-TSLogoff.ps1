<#
.SYNOPSIS
    Desktop Management Logoff Script - Terminal Server/Citrix
    
.DESCRIPTION
    Main entry point for Desktop Management Suite during user logoff on Terminal Server/Citrix.
    Uses workflow engine to execute steps defined in Config\Workflow-TSLogoff.psd1
    
.PARAMETER VerboseLogging
    Enable verbose logging
    
.PARAMETER MaxLogAge
    Maximum age of log files in days (default: 60)
    
.EXAMPLE
    .\DesktopManagement-TSLogoff.ps1
    .\DesktopManagement-TSLogoff.ps1 -VerboseLogging
    
.NOTES
    Version: 2.0.0
    Replacement for: VB/DesktopManagement.wsf //Job:GDP_TS_Logoff
    
    To add/remove/reorder actions:
    - Edit Config\Workflow-TSLogoff.psd1
    - No changes to this script needed!
#>

Param(
    [Parameter(Mandatory=$False)]
    [Switch]$VerboseLogging,
    
    [Parameter(Mandatory=$False)]
    [Int]$MaxLogAge = 60
)

# Script metadata
[String]$ScriptVersion = "5.0"
[String]$JobType = "TSLogoff"
[DateTime]$ScriptStartTime = Get-Date

$ErrorActionPreference = 'Continue'

# ============================================================================
# Import Framework Modules
# ============================================================================
Try {
    Import-Module "$PSScriptRoot\Modules\Framework\DMLogger.psm1" -Force -ErrorAction Stop
    Import-Module "$PSScriptRoot\Modules\Framework\DMCommon.psm1" -Force -ErrorAction Stop
    Import-Module "$PSScriptRoot\Modules\Framework\DMRegistry.psm1" -Force -ErrorAction Stop
    Import-Module "$PSScriptRoot\Modules\Utilities\Test-Environment.psm1" -Force -ErrorAction Stop
    Import-Module "$PSScriptRoot\Modules\Framework\DMComputer.psm1" -Force -ErrorAction Stop
    Import-Module "$PSScriptRoot\Modules\Framework\DMUser.psm1" -Force -ErrorAction Stop
    Import-Module "$PSScriptRoot\Modules\Framework\DMWorkflowEngine.psm1" -Force -ErrorAction Stop
    
    # Import service modules
    Import-Module "$PSScriptRoot\Modules\Services\DMServiceCommon.psm1" -Force -ErrorAction Stop
    Import-Module "$PSScriptRoot\Modules\Services\DMInventoryService.psm1" -Force -ErrorAction Stop
    
    Write-Host "✅ All modules imported successfully" -ForegroundColor Green
} Catch {
    Write-Error "Failed to import required modules: $($_.Exception.Message)"
    Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host "Execution Policy: $(Get-ExecutionPolicy)" -ForegroundColor Yellow
    Write-Host "Available Functions:" -ForegroundColor Yellow
    Try {
        Get-Command Test-DMComputerDomainJoined, Get-DMComputerInfo, Get-DMUserInfo -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Host "  ✅ $($_.Name)" -ForegroundColor Green
        }
    } Catch {
        Write-Host "  ❌ Function check failed" -ForegroundColor Red
    }
    Exit 1
}

# ============================================================================
# Initialize Logging
# ============================================================================
Try {
    [Boolean]$InitResult = Initialize-DMLog -JobType $JobType -Version $ScriptVersion -VerboseLogging:$VerboseLogging -MaxLogAge $MaxLogAge
    
    If (-not $InitResult) {
        Write-Error "Failed to initialize logging"
        Exit 1
    }
} Catch {
    Write-Error "Failed to initialize logging: $($_.Exception.Message)"
    Exit 1
}

Write-DMLog "========================================" -NoTimestamp
Write-DMLog "Desktop Management Suite - TS Logoff Script" -NoTimestamp
Write-DMLog "Version: $ScriptVersion" -NoTimestamp
Write-DMLog "========================================" -NoTimestamp
Write-DMLog ""

# ============================================================================
# Get Computer and User Information
# ============================================================================
Write-DMLog "Gathering computer and user information..."

Try {
    $Computer = Get-DMComputerInfo
    $User = Get-DMUserInfo
    
    If ($Null -eq $Computer -or $Null -eq $User) {
        Write-DMLog "Failed to gather computer or user information" -Level Error
        Set-DMLogError -ErrorLevel 2 -ErrorDescription "Failed to gather system information"
        Export-DMLog
        Exit 1
    }
    
    # Log detailed computer information (matching VBScript format)
    Write-DMLog "Computer Information:" -Level Info
    Write-DMLog "  Name: $($Computer.Name)" -Level Info
    Write-DMLog "  DN: $($Computer.DistinguishedName)" -Level Info
    Write-DMLog "  Domain: $($Computer.Domain)" -Level Info
    Write-DMLog "  ShortDomain: $($Computer.ShortDomain)" -Level Info
    Write-DMLog "  Site: $($Computer.Site)" -Level Info
    Write-DMLog "  OU: $($Computer.OUMapping)" -Level Info
    Write-DMLog "  CityCode: $($Computer.CityCode)" -Level Info
    Write-DMLog "  IPAddresses: $($Computer.IPAddresses -join ', ')" -Level Info
    Write-DMLog "  OSCaption: $($Computer.OSCaption)" -Level Info
    Write-DMLog "  DesktopOS: $($Computer.IsDesktop)" -Level Info
    Write-DMLog "  Vpnconnected: $($Computer.IsVPNConnected)" -Level Info
    Write-DMLog "  Groups: $($Computer.Groups.Count) groups" -Level Info
    If ($Computer.Groups.Count -gt 0) {
        $Computer.Groups | ForEach-Object { Write-DMLog "    $($_.Name)" -Level Info }
    }
    Write-DMLog ""
    
    # Log detailed user information (matching VBScript format)
    Write-DMLog "User Information:" -Level Info
    Write-DMLog "  Name: $($User.Name)" -Level Info
    Write-DMLog "  DN: $($User.DistinguishedName)" -Level Info
    Write-DMLog "  Domain: $($User.Domain)" -Level Info
    Write-DMLog "  ShortDomain: $($User.ShortDomain)" -Level Info
    Write-DMLog "  LogonServer: $($User.LogonServer)" -Level Info
    Write-DMLog "  OU: $($User.OUMapping)" -Level Info
    Write-DMLog "  Session Type: $($User.SessionType)" -Level Info
    Write-DMLog "  Groups: $($User.Groups.Count) groups" -Level Info
    If ($User.Groups.Count -gt 0) {
        $User.Groups | ForEach-Object { Write-DMLog "    $($_.Name)" -Level Info }
    }
    Write-DMLog ""
} Catch {
    Write-DMLog "Error gathering system information: $($_.Exception.Message)" -Level Error
    Set-DMLogError -ErrorLevel 2 -ErrorDescription $_.Exception.Message
    Export-DMLog
    Exit 1
}

# ============================================================================
# Execute Workflow
# ============================================================================
Write-DMLog "Starting workflow execution..." -Level Info
Write-DMLog ""

Try {
    # Build workflow context
    [Hashtable]$Context = @{
        UserInfo = $User
        ComputerInfo = $Computer
        JobType = $JobType
        ScriptVersion = $ScriptVersion
    }
    
    # Load and execute workflow
    [String]$WorkflowFile = "$PSScriptRoot\Config\Workflow-$JobType.psd1"
    
    [Boolean]$WorkflowResult = Invoke-DMWorkflow -WorkflowFile $WorkflowFile -Context $Context
    
    If (-not $WorkflowResult) {
        Write-DMLog "Workflow execution failed" -Level Error
        Set-DMLogError -ErrorLevel 2 -ErrorDescription "Workflow execution failed"
    }
} Catch {
    Write-DMLog "Fatal error during workflow execution: $($_.Exception.Message)" -Level Error
    Set-DMLogError -ErrorLevel 2 -ErrorDescription $_.Exception.Message
}

# ============================================================================
# Finalize
# ============================================================================
Write-DMLog "========================================" -NoTimestamp
Write-DMLog "Desktop Management TS Logoff - Completed" -NoTimestamp
Write-DMLog "========================================" -NoTimestamp

Export-DMLog

# Write execution metadata to registry (matches VBScript format)
Try {
    [DateTime]$ScriptEndTime = Get-Date
    [Int]$RunTimeSeconds = [Math]::Round(($ScriptEndTime - $ScriptStartTime).TotalSeconds)
    [String]$LogPath = Get-DMLogPath
    
    Set-DMExecutionMetadata -JobType $JobType `
                             -Version $ScriptVersion `
                             -ScriptPath $PSCommandPath `
                             -LogFilePath $LogPath `
                             -StartTime $ScriptStartTime `
                             -EndTime $ScriptEndTime `
                             -RunTimeSeconds $RunTimeSeconds
} Catch {
    Write-DMLog "Warning: Could not write execution metadata to registry" -Level Warning
}

Exit 0
