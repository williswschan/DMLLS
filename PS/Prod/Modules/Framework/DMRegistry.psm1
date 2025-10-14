<#
.SYNOPSIS
    Desktop Management Registry Operations Module
    
.DESCRIPTION
    Provides registry read/write operations for the Desktop Management Suite.
    Handles execution tracking, version storage, and configuration persistence.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: Registry operations in VBScript modules
#>

<#
.SYNOPSIS
    Reads a registry value safely.
    
.DESCRIPTION
    Reads a registry value with error handling. Returns default value if not found.
    
.PARAMETER Path
    Registry path (e.g., "HKCU:\Software\Nomura\GDP")
    
.PARAMETER Name
    Value name
    
.PARAMETER DefaultValue
    Default value to return if not found
    
.OUTPUTS
    Object - registry value or default
    
.EXAMPLE
    $Value = Get-DMRegistryValue -Path "HKCU:\Software\Nomura\GDP" -Name "Version" -DefaultValue "0.0"
#>
Function Get-DMRegistryValue {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path,
        
        [Parameter(Mandatory=$True)]
        [String]$Name,
        
        [Parameter(Mandatory=$False)]
        [Object]$DefaultValue = $Null
    )
    
    Try {
        If (Test-Path -Path $Path) {
            [Object]$Value = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            If ($Null -ne $Value) {
                Return $Value.$Name
            }
        }
        Return $DefaultValue
    }
    Catch {
        Return $DefaultValue
    }
}

<#
.SYNOPSIS
    Writes a registry value safely.
    
.DESCRIPTION
    Writes a registry value, creating the path if necessary.
    
.PARAMETER Path
    Registry path
    
.PARAMETER Name
    Value name
    
.PARAMETER Value
    Value to write
    
.PARAMETER Type
    Value type (String, DWord, QWord, Binary, MultiString, ExpandString)
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    Set-DMRegistryValue -Path "HKCU:\Software\Nomura\GDP" -Name "Version" -Value "2.0" -Type String
#>
Function Set-DMRegistryValue {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path,
        
        [Parameter(Mandatory=$True)]
        [String]$Name,
        
        [Parameter(Mandatory=$True)]
        [Object]$Value,
        
        [Parameter(Mandatory=$False)]
        [ValidateSet('String', 'DWord', 'QWord', 'Binary', 'MultiString', 'ExpandString')]
        [String]$Type = 'String'
    )
    
    Try {
        # Create path if it doesn't exist
        If (-not (Test-Path -Path $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        
        # Check if value exists
        [Object]$Existing = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        
        If ($Null -eq $Existing) {
            # Create new value
            New-ItemProperty -Path $Path -Name $Name -Value $Value -PropertyType $Type -Force | Out-Null
        } Else {
            # Update existing value
            Set-ItemProperty -Path $Path -Name $Name -Value $Value -Force
        }
        
        Return $True
    }
    Catch {
        Write-Warning "Failed to write registry value '$Path\$Name': $($_.Exception.Message)"
        Return $False
    }
}

<#
.SYNOPSIS
    Tests if a registry path exists.
    
.DESCRIPTION
    Checks if a registry path exists.
    
.PARAMETER Path
    Registry path to test
    
.OUTPUTS
    Boolean - true if exists
    
.EXAMPLE
    $Exists = Test-DMRegistryPath "HKCU:\Software\Nomura\GDP"
#>
Function Test-DMRegistryPath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path
    )
    
    Return (Test-Path -Path $Path)
}

<#
.SYNOPSIS
    Deletes a registry value.
    
.DESCRIPTION
    Removes a registry value if it exists.
    
.PARAMETER Path
    Registry path
    
.PARAMETER Name
    Value name to delete
    
.OUTPUTS
    Boolean - true if successful or value didn't exist
    
.EXAMPLE
    Remove-DMRegistryValue -Path "HKCU:\Software\Nomura\GDP" -Name "OldSetting"
#>
Function Remove-DMRegistryValue {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path,
        
        [Parameter(Mandatory=$True)]
        [String]$Name
    )
    
    Try {
        If (Test-Path -Path $Path) {
            Remove-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
        }
        Return $True
    }
    Catch {
        Write-Warning "Failed to delete registry value '$Path\$Name': $($_.Exception.Message)"
        Return $False
    }
}

<#
.SYNOPSIS
    Writes Desktop Management execution metadata to registry.
    
.DESCRIPTION
    Records script execution information in registry for tracking.
    Matches VBScript pattern of storing execution data.
    
.PARAMETER JobType
    Job type (Logon, Logoff, TSLogon, TSLogoff)
    
.PARAMETER Version
    Script version
    
.PARAMETER UserName
    User name
    
.PARAMETER ComputerName
    Computer name
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    Set-DMExecutionMetadata -JobType "Logon" -Version "2.0" -UserName "jsmith" -ComputerName "WKS001"
#>
Function Set-DMExecutionMetadata {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$JobType,
        
        [Parameter(Mandatory=$True)]
        [String]$Version,
        
        [Parameter(Mandatory=$False)]
        [String]$ScriptPath = $PSCommandPath,
        
        [Parameter(Mandatory=$False)]
        [String]$LogFilePath = "",
        
        [Parameter(Mandatory=$False)]
        [DateTime]$StartTime = (Get-Date),
        
        [Parameter(Mandatory=$False)]
        [DateTime]$EndTime = (Get-Date),
        
        [Parameter(Mandatory=$False)]
        [Int]$RunTimeSeconds = 0
    )
    
    Try {
        [String]$BasePath = "HKCU:\Software\Nomura\GDP\Desktop Management"
        
        # Create base path if it doesn't exist
        If (-not (Test-Path -Path $BasePath)) {
            New-Item -Path $BasePath -Force | Out-Null
        }
        
        # Convert to WMI datetime format (matching VBScript)
        # Format: yyyyMMddHHmmss.ffffff+000
        [String]$StartTimeWMI = $StartTime.ToString("yyyyMMddHHmmss") + ".000000+000"
        [String]$EndTimeWMI = $EndTime.ToString("yyyyMMddHHmmss") + ".000000+000"
        
        # Write execution metadata - EXACTLY as VBScript does (flat structure)
        [Void](Set-DMRegistryValue -Path $BasePath -Name "$JobType - Script Name" -Value $ScriptPath -Type String)
        [Void](Set-DMRegistryValue -Path $BasePath -Name "$JobType - Log File" -Value $LogFilePath -Type String)
        [Void](Set-DMRegistryValue -Path $BasePath -Name "$JobType - Start Time" -Value $StartTimeWMI -Type String)
        [Void](Set-DMRegistryValue -Path $BasePath -Name "$JobType - End Time" -Value $EndTimeWMI -Type String)
        [Void](Set-DMRegistryValue -Path $BasePath -Name "$JobType - Run Time (seconds)" -Value $RunTimeSeconds.ToString() -Type String)
        [Void](Set-DMRegistryValue -Path $BasePath -Name "$JobType - Script Version" -Value $Version -Type String)
        
        Return $True
    }
    Catch {
        Write-Warning "Failed to write execution metadata: $($_.Exception.Message)"
        Return $False
    }
}

<#
.SYNOPSIS
    Gets Desktop Management execution metadata from registry.
    
.DESCRIPTION
    Retrieves last execution information from registry.
    
.PARAMETER JobType
    Job type to query
    
.OUTPUTS
    PSCustomObject with execution metadata
    
.EXAMPLE
    $Metadata = Get-DMExecutionMetadata -JobType "Logon"
#>
Function Get-DMExecutionMetadata {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$JobType
    )
    
    Try {
        [String]$BasePath = "HKCU:\Software\Nomura\GDP\Desktop Management\$JobType"
        
        If (-not (Test-Path -Path $BasePath)) {
            Return $Null
        }
        
        Return [PSCustomObject]@{
            PSTypeName = 'DM.ExecutionMetadata'
            Version = Get-DMRegistryValue -Path $BasePath -Name "Version" -DefaultValue ""
            LastRun = Get-DMRegistryValue -Path $BasePath -Name "LastRun" -DefaultValue ""
            UserName = Get-DMRegistryValue -Path $BasePath -Name "UserName" -DefaultValue ""
            ComputerName = Get-DMRegistryValue -Path $BasePath -Name "ComputerName" -DefaultValue ""
            ScriptEngine = Get-DMRegistryValue -Path $BasePath -Name "ScriptEngine" -DefaultValue ""
            PSVersion = Get-DMRegistryValue -Path $BasePath -Name "PSVersion" -DefaultValue ""
        }
    }
    Catch {
        Return $Null
    }
}

<#
.SYNOPSIS
    Reads binary registry value.
    
.DESCRIPTION
    Reads a binary registry value and returns byte array.
    Used for Outlook PST registry parsing.
    
.PARAMETER Path
    Registry path
    
.PARAMETER Name
    Value name
    
.OUTPUTS
    Byte array or $Null if not found
    
.EXAMPLE
    $Bytes = Get-DMRegistryBinaryValue -Path "HKCU:\Software\..." -Name "BinaryData"
#>
Function Get-DMRegistryBinaryValue {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path,
        
        [Parameter(Mandatory=$True)]
        [String]$Name
    )
    
    Try {
        If (Test-Path -Path $Path) {
            [Object]$Value = Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
            If ($Null -ne $Value) {
                Return $Value.$Name
            }
        }
        Return $Null
    }
    Catch {
        Return $Null
    }
}

<#
.SYNOPSIS
    Enumerates all subkeys in a registry path.
    
.DESCRIPTION
    Gets all subkey names under a registry path.
    
.PARAMETER Path
    Registry path
    
.OUTPUTS
    Array of subkey names
    
.EXAMPLE
    $Subkeys = Get-DMRegistrySubKeys -Path "HKCU:\Software\Nomura"
#>
Function Get-DMRegistrySubKeys {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path
    )
    
    Try {
        If (Test-Path -Path $Path) {
            [Object]$Key = Get-Item -Path $Path
            Return $Key.GetSubKeyNames()
        }
        Return @()
    }
    Catch {
        Return @()
    }
}

<#
.SYNOPSIS
    Gets all value names in a registry path.
    
.DESCRIPTION
    Returns all value names (properties) in a registry key.
    
.PARAMETER Path
    Registry path
    
.OUTPUTS
    Array of value names
    
.EXAMPLE
    $Values = Get-DMRegistryValueNames -Path "HKCU:\Software\Nomura\GDP"
#>
Function Get-DMRegistryValueNames {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path
    )
    
    Try {
        If (Test-Path -Path $Path) {
            [Object]$Item = Get-Item -Path $Path
            Return $Item.Property
        }
        Return @()
    }
    Catch {
        Return @()
    }
}

<#
.SYNOPSIS
    Imports a registry file.
    
.DESCRIPTION
    Imports a .reg file using regedit.
    Used for IE Zone configuration (legacy).
    
.PARAMETER FilePath
    Path to .reg file
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    Import-DMRegistryFile -FilePath "\\server\share\IEZones.reg"
#>
Function Import-DMRegistryFile {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$FilePath
    )
    
    Try {
        If (-not (Test-Path -Path $FilePath)) {
            Write-Warning "Registry file not found: $FilePath"
            Return $False
        }
        
        [Object]$Result = Start-Process -FilePath "regedit.exe" -ArgumentList "/s `"$FilePath`"" -Wait -NoNewWindow -PassThru
        
        Return ($Result.ExitCode -eq 0)
    }
    Catch {
        Write-Warning "Failed to import registry file: $($_.Exception.Message)"
        Return $False
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Get-DMRegistryValue',
    'Set-DMRegistryValue',
    'Test-DMRegistryPath',
    'Remove-DMRegistryValue',
    'Set-DMExecutionMetadata',
    'Get-DMExecutionMetadata',
    'Get-DMRegistryBinaryValue',
    'Get-DMRegistrySubKeys',
    'Get-DMRegistryValueNames',
    'Import-DMRegistryFile'
)

