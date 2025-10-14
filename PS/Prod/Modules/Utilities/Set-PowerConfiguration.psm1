<#
.SYNOPSIS
    Desktop Management Power Configuration Module
    
.DESCRIPTION
    Configures monitor timeout settings based on screensaver policy.
    Sets appropriate timeout for physical machines, disables for VMs.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: PowerCFG_W10.vbs
#>

# Import required modules
Using Module .\Test-Environment.psm1
Using Module ..\Framework\DMLogger.psm1
Using Module ..\Framework\DMRegistry.psm1

<#
.SYNOPSIS
    Configures power settings based on job type and environment.
    
.DESCRIPTION
    Sets monitor timeout based on screensaver GPO settings.
    
    Logic:
    - VMs: No timeout (0 minutes)
    - Logon: Timeout = (ScreenSaverTimeout + 300) / 60 minutes
    - Logoff: Revert to 20 minutes
    
.PARAMETER JobType
    Job type (Logon or Logoff)
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    Set-DMPowerConfiguration -JobType "Logon"
#>
Function Set-DMPowerConfiguration {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [ValidateSet('Logon', 'Logoff')]
        [String]$JobType
    )
    
    Try {
        Write-DMLog "PowerCFG: Initialized" -Level Info
        Write-DMLog "PowerCFG: JobType $($JobType.ToUpper())" -Level Info
        Write-DMLog "PowerCFG: About to set Power Scheme" -Level Info
        
        # Check if running on a virtual machine
        [Object]$VMInfo = Test-DMVirtualMachine
        
        If ($VMInfo.IsVirtual) {
            # Virtual machines - no monitor timeout
            Write-DMLog "PowerCFG: No monitor time-out for VM (Platform: $($VMInfo.Platform))" -Level Info
            
            [Boolean]$Success = Set-DMMonitorTimeout -TimeoutMinutes 0
            
            Write-DMLog "PowerCFG: Completed" -Level Verbose
            Return $Success
        }
        
        # Physical machines - calculate timeout based on job type
        If ($JobType -eq "Logon") {
            # Get screensaver timeout from GPO
            [String]$GPOPath = "HKCU:\Software\Policies\Microsoft\Windows\Control Panel\Desktop"
            [String]$ScreenSaverTimeout = Get-DMRegistryValue -Path $GPOPath -Name "ScreenSaveTimeOut" -DefaultValue "0"
            
            # Calculate monitor timeout
            [Int]$TimeoutSeconds = 0
            If (-not [String]::IsNullOrEmpty($ScreenSaverTimeout)) {
                Try {
                    $TimeoutSeconds = [Int]$ScreenSaverTimeout
                } Catch {
                    $TimeoutSeconds = 0
                }
            }
            
            # Calculate: (ScreenSaverTimeout + 300) / 60 minutes
            [Int]$TimeoutMinutes = 0
            If ($TimeoutSeconds -gt 0) {
                $TimeoutMinutes = [Math]::Floor(($TimeoutSeconds + 300) / 60)
            }
            
            Write-DMLog "PowerCFG: Effective monitor time-out value is $TimeoutMinutes minutes" -Level Info
            
            [Boolean]$Success = Set-DMMonitorTimeout -TimeoutMinutes $TimeoutMinutes
            
        } Else {
            # Logoff - revert to 20 minutes
            Write-DMLog "PowerCFG: Effective monitor time-out value revert to 20 mins" -Level Info
            
            [Boolean]$Success = Set-DMMonitorTimeout -TimeoutMinutes 20
        }
        
        Write-DMLog "PowerCFG: Completed" -Level Verbose
        Return $Success
    }
    Catch {
        Write-DMLog "PowerCFG: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Sets monitor timeout using PowerCFG.exe.
    
.DESCRIPTION
    Executes PowerCFG.exe to set monitor timeout on AC power.
    
.PARAMETER TimeoutMinutes
    Timeout in minutes (0 = never)
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    Set-DMMonitorTimeout -TimeoutMinutes 15
#>
Function Set-DMMonitorTimeout {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Int]$TimeoutMinutes
    )
    
    Try {
        [String]$Command = "PowerCFG.exe"
        [String]$Arguments = "-change -monitor-timeout-ac $TimeoutMinutes"
        
        Write-DMLog "PowerCFG: About to launch command '$Command $Arguments'" -Level Info
        
        # Execute PowerCFG
        [Object]$Process = Start-Process -FilePath $Command -ArgumentList $Arguments -Wait -NoNewWindow -PassThru
        
        [Int]$ExitCode = $Process.ExitCode
        
        If ($ExitCode -eq 0) {
            Write-DMLog "PowerCFG: Command run successfully" -Level Info
            Return $True
        } Else {
            Write-DMLog "PowerCFG: Command completed with error(s). Exit code '$ExitCode'" -Level Warning
            Return $False
        }
    }
    Catch {
        Write-DMLog "PowerCFG: Error running command: $($_.Exception.Message)" -Level Error
        Return $False
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Set-DMPowerConfiguration',
    'Set-DMMonitorTimeout'
)

