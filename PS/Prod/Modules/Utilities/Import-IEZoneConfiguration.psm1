<#
.SYNOPSIS
    Desktop Management IE Zone Configuration Module (LEGACY)
    
.DESCRIPTION
    Imports Internet Explorer zone configuration from registry files.
    This is a LEGACY module for IE zone management.
    
    NOTE: This module is deprecated and disabled by default (IE is deprecated).
    Only enable if explicitly required for legacy browser support.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: LegacyModules/ManageIEZones.vbs
#>

# Import required modules
Using Module ..\Framework\DMLogger.psm1
Using Module ..\Framework\DMRegistry.psm1
Using Module ..\Services\DMServiceCommon.psm1

<#
.SYNOPSIS
    Imports IE zone configuration from registry file.
    
.DESCRIPTION
    Imports IE zone settings from network share based on:
    - Job type (Logon vs Startup)
    - Pilot vs Production group membership
    - Domain/Region
    
    Registry File Path Format:
    \\<Domain>\Apps\ConfigFiles\IEZones\<Pilot|Prod>\IEZones-<U|M>.reg
    
.PARAMETER JobType
    Job type (Logon or Startup)
    
.PARAMETER UserInfo
    User information (for Logon jobs)
    
.PARAMETER ComputerInfo
    Computer information (for Startup jobs)
    
.OUTPUTS
    Boolean - true if successful or not applicable
    
.EXAMPLE
    Import-DMIEZoneConfiguration -JobType "Logon" -UserInfo $User
#>
Function Import-DMIEZoneConfiguration {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [ValidateSet('Logon', 'Startup')]
        [String]$JobType,
        
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$UserInfo = $Null,
        
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$ComputerInfo = $Null
    )
    
    Try {
        Write-DMLog "Manage IE Zones: Manage IE zone configuration and assign sites and domains to zones" -Level Info
        
        [String]$RegFileServerPath = ""
        [String]$RegFileName = ""
        [String]$Domain = ""
        [String]$IEZoneFolder = "Prod\"
        
        If ($JobType -eq "Logon") {
            # User-based (Logon)
            If ($Null -eq $UserInfo) {
                Write-DMLog "Manage IE Zones: UserInfo required for Logon job type" -Level Warning
                Return $False
            }
            
            $Domain = $UserInfo.Domain
            $RegFileName = "IEZones-U.reg"
            
            # Determine Pilot vs Prod based on group membership
            If (Test-DMUserPartOfGroup -UserInfo $UserInfo -GroupName "Pilot Desktop Management Script") {
                $IEZoneFolder = "Pilot\"
            }
            
        } Else {
            # Computer-based (Startup)
            If ($Null -eq $ComputerInfo) {
                Write-DMLog "Manage IE Zones: ComputerInfo required for Startup job type" -Level Warning
                Return $False
            }
            
            $Domain = $ComputerInfo.Domain
            $RegFileName = "IEZones-M.reg"
            
            # Determine Pilot vs Prod based on computer group membership
            If (Test-DMHostPartOfGroup -ComputerInfo $ComputerInfo -GroupName "Pilot Desktop Management Script") {
                $IEZoneFolder = "Pilot\"
            }
        }
        
        # Validate domain is not empty
        If ([String]::IsNullOrEmpty($Domain)) {
            Write-DMLog "Manage IE Zones: Domain is empty (not domain-joined), cannot determine registry file path" -Level Verbose
            Write-DMLog "Manage IE Zones: Completed" -Level Info
            Return $True  # Return true as this is expected for non-domain computers
        }
        
        # Build registry file path
        $RegFileServerPath = "\\$Domain\Apps\ConfigFiles\IEZones\$IEZoneFolder$RegFileName"
        
        Write-DMLog "Manage IE Zones: Reg merge file path is '$RegFileServerPath'" -Level Info
        
        # Check if file exists
        If (-not (Test-Path -Path $RegFileServerPath)) {
            Write-DMLog "Manage IE Zones: Cannot access the '$RegFileServerPath' file" -Level Warning
            Write-DMLog "Manage IE Zones: Completed" -Level Info
            Return $False
        }
        
        # Copy to local temp for better performance
        [String]$TempPath = Join-Path $env:TEMP $RegFileName
        
        Try {
            Copy-Item -Path $RegFileServerPath -Destination $TempPath -Force
            Write-DMLog "Manage IE Zones: Reg merge file copied to '$TempPath'" -Level Verbose
            [String]$RegFilePath = $TempPath
        } Catch {
            # Use server path if copy fails
            Write-DMLog "Manage IE Zones: Could not copy to temp, using server path" -Level Verbose
            [String]$RegFilePath = $RegFileServerPath
        }
        
        # Import registry file
        Write-DMLog "Manage IE Zones: About to import reg file: $RegFilePath" -Level Verbose
        
        [Boolean]$Success = Import-DMRegistryFile -FilePath $RegFilePath
        
        If ($Success) {
            Write-DMLog "Manage IE Zones: Completed import successfully" -Level Info
        } Else {
            Write-DMLog "Manage IE Zones: Import completed with errors" -Level Warning
        }
        
        Write-DMLog "Manage IE Zones: Completed" -Level Info
        Return $Success
    }
    Catch {
        Write-DMLog "Manage IE Zones: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Import-DMIEZoneConfiguration'
)

