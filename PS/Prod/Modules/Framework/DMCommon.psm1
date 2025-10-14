<#
.SYNOPSIS
    Desktop Management Common Utilities Module
    
.DESCRIPTION
    Provides common utility functions used across the Desktop Management Suite.
    Includes XML escaping, string manipulation, and helper functions.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: Common functions from VBScript modules
#>

<#
.SYNOPSIS
    Escapes XML special characters in text.
    
.DESCRIPTION
    Replaces XML special characters with entities for safe SOAP/XML usage.
    Matches VBScript EscapeXMLText() function.
    
.PARAMETER Text
    Text to escape
    
.OUTPUTS
    String - escaped text
    
.EXAMPLE
    $SafeText = ConvertTo-DMXMLSafeText "Price < 100 & Quantity > 5"
    # Returns: "Price &lt; 100 &amp; Quantity &gt; 5"
#>
Function ConvertTo-DMXMLSafeText {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True, ValueFromPipeline=$True)]
        [AllowEmptyString()]
        [String]$Text
    )
    
    Process {
        If ([String]::IsNullOrEmpty($Text)) {
            Return ""
        }
        
        [String]$Result = $Text
        $Result = $Result.Replace("&", "&amp;")   # Must be first
        $Result = $Result.Replace("<", "&lt;")
        $Result = $Result.Replace(">", "&gt;")
        $Result = $Result.Replace('"', "&quot;")
        $Result = $Result.Replace("'", "&apos;")
        
        Return $Result
    }
}

<#
.SYNOPSIS
    Converts wildcard pattern to regex pattern.
    
.DESCRIPTION
    Converts wildcard patterns (* and ?) to regex for matching.
    Used for disconnect patterns in drive/printer mapping.
    
.PARAMETER Pattern
    Wildcard pattern
    
.OUTPUTS
    String - regex pattern
    
.EXAMPLE
    $Regex = ConvertTo-DMRegexPattern "H:\Users\*\Documents"
    # Returns: "H:\\Users\\.*\\Documents"
#>
Function ConvertTo-DMRegexPattern {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Pattern
    )
    
    [String]$Result = [Regex]::Escape($Pattern)
    $Result = $Result.Replace("\*", ".*")  # * = any characters
    $Result = $Result.Replace("\?", ".")   # ? = single character
    
    Return "^$Result$"  # Exact match
}

<#
.SYNOPSIS
    Tests if a string matches a wildcard pattern.
    
.DESCRIPTION
    Matches string against wildcard pattern (supports * and ?).
    
.PARAMETER Text
    Text to test
    
.PARAMETER Pattern
    Wildcard pattern
    
.OUTPUTS
    Boolean - true if matches
    
.EXAMPLE
    $Matches = Test-DMWildcardMatch -Text "\\server\share\data" -Pattern "\\server\*"
#>
Function Test-DMWildcardMatch {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Text,
        
        [Parameter(Mandatory=$True)]
        [String]$Pattern
    )
    
    [String]$RegexPattern = ConvertTo-DMRegexPattern -Pattern $Pattern
    Return $Text -match $RegexPattern
}

<#
.SYNOPSIS
    Expands environment variables in a path.
    
.DESCRIPTION
    Expands %VARIABLE% style environment variables in paths.
    Supports nested paths and UNC paths.
    
.PARAMETER Path
    Path with environment variables
    
.OUTPUTS
    String - expanded path
    
.EXAMPLE
    $Path = Expand-DMEnvironmentPath "%HOMEDRIVE%\%HOMEPATH%\Documents"
#>
Function Expand-DMEnvironmentPath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True, ValueFromPipeline=$True)]
        [String]$Path
    )
    
    Process {
        Return [System.Environment]::ExpandEnvironmentVariables($Path)
    }
}

<#
.SYNOPSIS
    Converts UNC path to drive letter if mapped.
    
.DESCRIPTION
    Finds the drive letter for a given UNC path if it's currently mapped.
    
.PARAMETER UNCPath
    UNC path to check
    
.OUTPUTS
    String - drive letter (e.g., "H:") or empty string if not mapped
    
.EXAMPLE
    $DriveLetter = ConvertFrom-DMUNCPath "\\server\share"
#>
Function ConvertFrom-DMUNCPath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$UNCPath
    )
    
    Try {
        [Array]$Drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -eq $UNCPath }
        
        If ($Drives.Count -gt 0) {
            Return "$($Drives[0].Name):"
        }
        
        Return ""
    }
    Catch {
        Return ""
    }
}

<#
.SYNOPSIS
    Converts drive letter to UNC path.
    
.DESCRIPTION
    Returns the UNC path for a mapped drive letter.
    
.PARAMETER DriveLetter
    Drive letter (e.g., "H:" or "H")
    
.OUTPUTS
    String - UNC path or empty string if not a network drive
    
.EXAMPLE
    $UNCPath = ConvertTo-DMUNCPath "H:"
#>
Function ConvertTo-DMUNCPath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DriveLetter
    )
    
    Try {
        # Normalize drive letter (remove colon if present)
        $DriveLetter = $DriveLetter.TrimEnd(':')
        
        [Object]$Drive = Get-PSDrive -Name $DriveLetter -PSProvider FileSystem -ErrorAction SilentlyContinue
        
        If ($Null -ne $Drive -and $Null -ne $Drive.DisplayRoot) {
            Return $Drive.DisplayRoot
        }
        
        Return ""
    }
    Catch {
        Return ""
    }
}

<#
.SYNOPSIS
    Pings a server to test connectivity.
    
.DESCRIPTION
    Tests network connectivity to a server using ICMP ping.
    
.PARAMETER ComputerName
    Computer or server name to ping
    
.PARAMETER Count
    Number of ping attempts (default: 1)
    
.PARAMETER Timeout
    Timeout in milliseconds (default: 1000)
    
.OUTPUTS
    Boolean - true if ping successful
    
.EXAMPLE
    $IsOnline = Test-DMServerPing -ComputerName "gdpmappercb.nomura.com"
#>
Function Test-DMServerPing {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$ComputerName,
        
        [Parameter(Mandatory=$False)]
        [Int]$Count = 1,
        
        [Parameter(Mandatory=$False)]
        [Int]$Timeout = 1000
    )
    
    Try {
        [Object]$PingResult = Test-Connection -ComputerName $ComputerName -Count $Count -Quiet -ErrorAction SilentlyContinue
        Return $PingResult
    }
    Catch {
        Return $False
    }
}

<#
.SYNOPSIS
    Gets file size in bytes.
    
.DESCRIPTION
    Returns the size of a file in bytes.
    
.PARAMETER Path
    File path
    
.OUTPUTS
    Int64 - file size in bytes, or 0 if file doesn't exist
    
.EXAMPLE
    $Size = Get-DMFileSize "C:\Data\file.pst"
#>
Function Get-DMFileSize {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path
    )
    
    Try {
        If (Test-Path -Path $Path -PathType Leaf) {
            [Object]$File = Get-Item -Path $Path
            Return $File.Length
        }
        Return 0
    }
    Catch {
        Return 0
    }
}

<#
.SYNOPSIS
    Gets file last modified date.
    
.DESCRIPTION
    Returns the last write time of a file.
    
.PARAMETER Path
    File path
    
.OUTPUTS
    DateTime - last write time, or $Null if file doesn't exist
    
.EXAMPLE
    $LastModified = Get-DMFileLastModified "C:\Data\file.pst"
#>
Function Get-DMFileLastModified {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path
    )
    
    Try {
        If (Test-Path -Path $Path -PathType Leaf) {
            [Object]$File = Get-Item -Path $Path
            Return $File.LastWriteTime
        }
        Return $Null
    }
    Catch {
        Return $Null
    }
}

<#
.SYNOPSIS
    Executes a command and returns output.
    
.DESCRIPTION
    Executes a command line and captures output.
    Used for legacy command-line utilities.
    
.PARAMETER Command
    Command to execute
    
.PARAMETER Arguments
    Command arguments
    
.OUTPUTS
    PSCustomObject with ExitCode, Output, and Error properties
    
.EXAMPLE
    $Result = Invoke-DMCommand -Command "whoami" -Arguments "/groups"
#>
Function Invoke-DMCommand {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Command,
        
        [Parameter(Mandatory=$False)]
        [String]$Arguments = ""
    )
    
    Try {
        [Object]$ProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
        $ProcessInfo.FileName = $Command
        $ProcessInfo.Arguments = $Arguments
        $ProcessInfo.RedirectStandardOutput = $True
        $ProcessInfo.RedirectStandardError = $True
        $ProcessInfo.UseShellExecute = $False
        $ProcessInfo.CreateNoWindow = $True
        
        [Object]$Process = New-Object System.Diagnostics.Process
        $Process.StartInfo = $ProcessInfo
        $Process.Start() | Out-Null
        
        [String]$Output = $Process.StandardOutput.ReadToEnd()
        [String]$ErrorOutput = $Process.StandardError.ReadToEnd()
        
        $Process.WaitForExit()
        [Int]$ExitCode = $Process.ExitCode
        
        Return [PSCustomObject]@{
            PSTypeName = 'DM.CommandResult'
            ExitCode = $ExitCode
            Output = $Output
            Error = $ErrorOutput
            Success = ($ExitCode -eq 0)
        }
    }
    Catch {
        Return [PSCustomObject]@{
            PSTypeName = 'DM.CommandResult'
            ExitCode = -1
            Output = ""
            Error = $_.Exception.Message
            Success = $False
        }
    }
}

<#
.SYNOPSIS
    Creates a COM object safely.
    
.DESCRIPTION
    Creates a COM object with error handling.
    
.PARAMETER ProgId
    COM ProgID (e.g., "Outlook.Application")
    
.OUTPUTS
    COM object or $Null if creation fails
    
.EXAMPLE
    $Outlook = New-DMCOMObject -ProgId "Outlook.Application"
#>
Function New-DMCOMObject {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$ProgId
    )
    
    Try {
        Return New-Object -ComObject $ProgId
    }
    Catch {
        Write-Warning "Failed to create COM object '$ProgId': $($_.Exception.Message)"
        Return $Null
    }
}

<#
.SYNOPSIS
    Splits a distinguished name into components.
    
.DESCRIPTION
    Parses an LDAP Distinguished Name into its component parts.
    
.PARAMETER DistinguishedName
    LDAP DN to parse
    
.OUTPUTS
    Array of hashtables with Type and Value properties
    
.EXAMPLE
    $Components = Split-DMDistinguishedName "CN=User,OU=USERS,OU=RESOURCES,OU=NYC,DC=nomura,DC=com"
#>
Function Split-DMDistinguishedName {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DistinguishedName
    )
    
    [Array]$Components = @()
    [Array]$Parts = $DistinguishedName -split ','
    
    ForEach ($Part in $Parts) {
        If ($Part -match '^(\w+)=(.+)$') {
            $Components += @{
                Type = $Matches[1]
                Value = $Matches[2]
            }
        }
    }
    
    Return $Components
}

# Export module members
Export-ModuleMember -Function @(
    'ConvertTo-DMXMLSafeText',
    'ConvertTo-DMRegexPattern',
    'Test-DMWildcardMatch',
    'Expand-DMEnvironmentPath',
    'ConvertFrom-DMUNCPath',
    'ConvertTo-DMUNCPath',
    'Test-DMServerPing',
    'Get-DMFileSize',
    'Get-DMFileLastModified',
    'Invoke-DMCommand',
    'New-DMCOMObject',
    'Split-DMDistinguishedName'
)

