<#
.SYNOPSIS
    Desktop Management Logging Module
    
.DESCRIPTION
    Provides centralized logging functionality for the Desktop Management Suite.
    Replaces VBScript LoggingObject class with PowerShell native logging.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: VB/Source/Main/Logging.vbs
#>

# Module-level variables
$Script:LogBuffer = New-Object System.Collections.ArrayList
$Script:LogFilePath = $Null
$Script:IsVerbose = $False
$Script:ErrorLevel = 0
$Script:ErrorDescription = ""

<#
.SYNOPSIS
    Initializes the logging system for a job session.
    
.DESCRIPTION
    Creates the log file path, initializes the buffer, and sets verbose mode.
    Log path format: %USERPROFILE%\Nomura\GDP\Desktop Management\<JobType>_<Computer>_<Timestamp>.log
    
.PARAMETER JobType
    Type of job (Logon, Logoff, TSLogon, TSLogoff)
    
.PARAMETER ComputerName
    Computer name (defaults to current computer)
    
.PARAMETER Verbose
    Enable verbose logging
    
.PARAMETER MaxLogAge
    Maximum age of log files in days (default: 60)
    
.EXAMPLE
    Initialize-DMLog -JobType "Logon" -Verbose
#>
Function Initialize-DMLog {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$JobType,
        
        [Parameter(Mandatory=$False)]
        [String]$Version = "1.1",
        
        [Parameter(Mandatory=$False)]
        [String]$ComputerName = $env:COMPUTERNAME,
        
        [Parameter(Mandatory=$False)]
        [Switch]$VerboseLogging,
        
        [Parameter(Mandatory=$False)]
        [Int]$MaxLogAge = 60
    )
    
    Try {
        # Set verbose mode
        $Script:IsVerbose = $VerboseLogging.IsPresent
        
        # Generate timestamp
        [String]$Timestamp = Get-Date -Format 'yyyyMMddHHmmss'
        
        # Build log file path
        [String]$BasePath = Join-Path $env:USERPROFILE "Nomura\GDP\Desktop Management"
        [String]$FileName = "${JobType}_${ComputerName}_${Timestamp}.log"
        $Script:LogFilePath = Join-Path $BasePath $FileName
        
        # Create directory if it doesn't exist
        [String]$LogDirectory = Split-Path -Path $Script:LogFilePath -Parent
        If (-not (Test-Path -Path $LogDirectory)) {
            New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
        }
        
        # Clear buffer
        $Script:LogBuffer.Clear()
        
        # Write initial header
        Write-DMLog "========================================" -NoTimestamp
        Write-DMLog "Desktop Management Suite - Version $Version" -NoTimestamp
        Write-DMLog "Job Type: $JobType" -NoTimestamp
        Write-DMLog "Computer: $ComputerName" -NoTimestamp
        Write-DMLog "User: $env:USERNAME" -NoTimestamp
        Write-DMLog "Domain: $env:USERDOMAIN" -NoTimestamp
        Write-DMLog "Start Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -NoTimestamp
        Write-DMLog "Verbose Mode: $($Script:IsVerbose)" -NoTimestamp
        Write-DMLog "========================================" -NoTimestamp
        Write-DMLog ""
        
        # Purge old log files
        If ($MaxLogAge -gt 0) {
            Remove-DMOldLogs -Path $BasePath -MaxAge $MaxLogAge
        }
        
        Return $True
    }
    Catch {
        Write-Error "Failed to initialize logging: $($_.Exception.Message)"
        Return $False
    }
}

<#
.SYNOPSIS
    Writes a message to the log.
    
.DESCRIPTION
    Adds timestamped message to log buffer. Respects verbose mode.
    
.PARAMETER Message
    Message to log
    
.PARAMETER Level
    Log level (Info, Warning, Error, Verbose)
    
.PARAMETER NoTimestamp
    Suppress timestamp prefix
    
.EXAMPLE
    Write-DMLog "Drive mapping completed successfully"
    Write-DMLog "VPN detected - skipping inventory" -Level Warning
#>
Function Write-DMLog {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True, Position=0, ValueFromPipeline=$True)]
        [AllowEmptyString()]
        [String]$Message,
        
        [Parameter(Mandatory=$False)]
        [ValidateSet('Info', 'Warning', 'Error', 'Verbose')]
        [String]$Level = 'Info',
        
        [Parameter(Mandatory=$False)]
        [Switch]$NoTimestamp
    )
    
    Process {
        # Skip verbose messages if not in verbose mode
        If ($Level -eq 'Verbose' -and -not $Script:IsVerbose) {
            Return
        }
        
        # Build log entry
        [String]$LogEntry = ""
        
        If (-not $NoTimestamp) {
            [String]$Timestamp = Get-Date -Format 'yyyyMMddHHmmss'
            $LogEntry = "[$Timestamp]"
        }
        
        # Add level prefix for non-Info messages
        Switch ($Level) {
            'Warning' { $LogEntry += "[WARNING]" }
            'Error'   { $LogEntry += "[ERROR]" }
            'Verbose' { $LogEntry += "[VERBOSE]" }
        }
        
        # Add message
        If ($LogEntry -ne "") {
            $LogEntry += " $Message"
        } Else {
            $LogEntry = $Message
        }
        
        # Add to buffer
        $Script:LogBuffer.Add($LogEntry) | Out-Null
        
        # Also write to console based on level
        Switch ($Level) {
            'Info'    { Write-Host $LogEntry }
            'Warning' { Write-Warning $LogEntry }
            'Error'   { Write-Error $LogEntry }
            'Verbose' { Write-Verbose $LogEntry }
        }
    }
}

<#
.SYNOPSIS
    Sets error information in the log.
    
.DESCRIPTION
    Records error level and description for tracking. Compatible with VBScript pattern.
    
.PARAMETER ErrorLevel
    Error severity level (0 = no error, 1 = warning, 2 = error)
    
.PARAMETER ErrorDescription
    Error description text
    
.EXAMPLE
    Set-DMLogError -ErrorLevel 2 -ErrorDescription "Failed to connect to mapper service"
#>
Function Set-DMLogError {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Int]$ErrorLevel,
        
        [Parameter(Mandatory=$True)]
        [String]$ErrorDescription
    )
    
    $Script:ErrorLevel = $ErrorLevel
    $Script:ErrorDescription = $ErrorDescription
    
    If ($ErrorLevel -gt 0) {
        [String]$LevelText = If ($ErrorLevel -eq 1) { "Warning" } Else { "Error" }
        Write-DMLog "Error Level $ErrorLevel - $ErrorDescription" -Level $LevelText
    }
}

<#
.SYNOPSIS
    Gets the current error information.
    
.DESCRIPTION
    Returns error level and description as PSCustomObject.
    
.OUTPUTS
    PSCustomObject with ErrorLevel and ErrorDescription properties
    
.EXAMPLE
    $ErrorInfo = Get-DMLogError
    If ($ErrorInfo.ErrorLevel -gt 0) { ... }
#>
Function Get-DMLogError {
    [CmdletBinding()]
    Param()
    
    Return [PSCustomObject]@{
        PSTypeName = 'DM.LogError'
        ErrorLevel = $Script:ErrorLevel
        ErrorDescription = $Script:ErrorDescription
    }
}

<#
.SYNOPSIS
    Exports the log buffer to file.
    
.DESCRIPTION
    Writes all buffered log entries to the log file and clears the buffer.
    
.PARAMETER Flush
    Flush buffer even if file write fails
    
.EXAMPLE
    Export-DMLog
#>
Function Export-DMLog {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [Switch]$Flush
    )
    
    Try {
        If ($Null -eq $Script:LogFilePath) {
            Write-Warning "Log not initialized. Call Initialize-DMLog first."
            Return $False
        }
        
        If ($Script:LogBuffer.Count -eq 0) {
            Write-Verbose "Log buffer is empty, nothing to export."
            Return $True
        }
        
        # Write footer
        Write-DMLog "========================================" -NoTimestamp
        Write-DMLog "End Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -NoTimestamp
        Write-DMLog "Error Level: $Script:ErrorLevel" -NoTimestamp
        If ($Script:ErrorLevel -gt 0) {
            Write-DMLog "Error Description: $Script:ErrorDescription" -NoTimestamp
        }
        Write-DMLog "========================================" -NoTimestamp
        
        # Write buffer to file
        $Script:LogBuffer | Out-File -FilePath $Script:LogFilePath -Encoding UTF8 -Append
        
        Write-Host "Log exported to: $Script:LogFilePath"
        
        # Clear buffer
        If ($Flush) {
            $Script:LogBuffer.Clear()
        }
        
        Return $True
    }
    Catch {
        Write-Error "Failed to export log: $($_.Exception.Message)"
        Return $False
    }
}

<#
.SYNOPSIS
    Removes old log files based on age.
    
.DESCRIPTION
    Deletes log files older than specified age in days.
    Matches VBScript purge functionality.
    
.PARAMETER Path
    Directory containing log files
    
.PARAMETER MaxAge
    Maximum age in days
    
.EXAMPLE
    Remove-DMOldLogs -Path "C:\Users\jsmith\Nomura\GDP\Desktop Management" -MaxAge 60
#>
Function Remove-DMOldLogs {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path,
        
        [Parameter(Mandatory=$True)]
        [Int]$MaxAge
    )
    
    Try {
        If (-not (Test-Path -Path $Path)) {
            Write-DMLog "Log directory does not exist: $Path" -Level Verbose
            Return
        }
        
        [DateTime]$CutoffDate = (Get-Date).AddDays(-$MaxAge)
        
        [Array]$OldLogs = Get-ChildItem -Path $Path -Filter "*.log" -Recurse -File | 
            Where-Object { $_.LastWriteTime -lt $CutoffDate }
        
        If ($OldLogs.Count -gt 0) {
            Write-DMLog "Purging $($OldLogs.Count) log file(s) older than $MaxAge days" -Level Verbose
            
            ForEach ($LogFile in $OldLogs) {
                Try {
                    Remove-Item -Path $LogFile.FullName -Force
                    Write-DMLog "Deleted: $($LogFile.Name)" -Level Verbose
                }
                Catch {
                    Write-DMLog "Failed to delete $($LogFile.Name): $($_.Exception.Message)" -Level Warning
                }
            }
        } Else {
            Write-DMLog "No log files older than $MaxAge days found" -Level Verbose
        }
    }
    Catch {
        Write-DMLog "Error purging old logs: $($_.Exception.Message)" -Level Warning
    }
}

<#
.SYNOPSIS
    Gets the current log file path.
    
.DESCRIPTION
    Returns the full path to the current log file.
    
.OUTPUTS
    String - log file path
    
.EXAMPLE
    $LogPath = Get-DMLogPath
#>
Function Get-DMLogPath {
    [CmdletBinding()]
    Param()
    
    Return $Script:LogFilePath
}

# Export module members
Export-ModuleMember -Function @(
    'Initialize-DMLog',
    'Write-DMLog',
    'Export-DMLog',
    'Set-DMLogError',
    'Get-DMLogError',
    'Remove-DMOldLogs',
    'Get-DMLogPath'
)

