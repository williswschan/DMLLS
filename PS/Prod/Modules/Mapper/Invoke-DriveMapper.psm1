<#
.SYNOPSIS
    Desktop Management Drive Mapper Module
    
.DESCRIPTION
    Maps network drives based on backend configuration.
    Handles conflict resolution, disconnect patterns, and home drive mapping.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: MapperDrives_W10.vbs
#>

# Import required modules
Using Module ..\Services\DMMapperService.psm1
Using Module ..\Framework\DMLogger.psm1
Using Module ..\Framework\DMCommon.psm1
Using Module ..\Framework\DMRegistry.psm1
Using Module ..\Utilities\Test-Environment.psm1
Using Module ..\Services\DMServiceCommon.psm1

<#
.SYNOPSIS
    Maps network drives based on backend configuration.
    
.DESCRIPTION
    Retrieves drive mappings from backend and applies them to the user's profile.
    Handles conflict resolution, disconnect patterns, VPN scenarios.
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Invoke-DMDriveMapper -UserInfo $User -ComputerInfo $Computer
#>
Function Invoke-DMDriveMapper {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo
    )
    
    Try {
        Write-DMLog "Mapper Drive: Retrieve drives and paths from service and map in user profile" -Level Info
        
        # Check if offline laptop (skip mapping)
        If (Test-DMUserPartOfGroup -UserInfo $UserInfo -GroupName "Laptop Offline PC") {
            Write-DMLog "Mapper Drive: User is part of 'Laptop Offline PC' group, skipping drive mapping" -Level Info
            Return $True
        }
        
        # Get current mapped drives
        [Array]$CurrentDrives = Get-DMMappedDrives -UserInfo $UserInfo
        Write-DMLog "Mapper Drive: Currently $($CurrentDrives.Count) drive(s) mapped" -Level Verbose
        
        # Get drive mappings from backend
        [Array]$DriveMappings = Get-DMDriveMappings -UserInfo $UserInfo -ComputerInfo $ComputerInfo
        
        If ($DriveMappings.Count -eq 0) {
            Write-DMLog "Mapper Drive: No drive mappings returned from service" -Level Info
            Return $True
        }
        
        Write-DMLog "Mapper Drive: Retrieved $($DriveMappings.Count) drive mapping(s) from service" -Level Info
        
        [Int]$SuccessfulMappings = 0
        [Int]$FailedMappings = 0
        [Int]$SuccessfulDisconnects = 0
        [Int]$FailedDisconnects = 0
        
        # Process disconnect patterns first
        ForEach ($Mapping in $DriveMappings) {
            If ($Mapping.DisconnectOnLogin) {
                [Boolean]$Success = Remove-DMDriveMappingByPattern -Pattern $Mapping.UncPath
                If ($Success) { $SuccessfulDisconnects++ } Else { $FailedDisconnects++ }
            }
        }
        
        # Process drive mappings
        ForEach ($Mapping in $DriveMappings) {
            If (-not $Mapping.DisconnectOnLogin) {
                [Boolean]$Success = Set-DMDriveMapping -Mapping $Mapping -CurrentDrives $CurrentDrives -UserInfo $UserInfo
                If ($Success) { $SuccessfulMappings++ } Else { $FailedMappings++ }
            }
        }
        
        # Map home drive (H:) if in allowed regions
        [Boolean]$HomeDriveMapped = Set-DMHomeDriveMapping -CurrentDrives $CurrentDrives
        
        Write-DMLog "Mapper Drive: Summary - Mapped: $SuccessfulMappings, Failed: $FailedMappings, Disconnected: $SuccessfulDisconnects" -Level Info
        
        Return ($FailedMappings -eq 0)
    }
    Catch {
        Write-DMLog "Mapper Drive: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Maps a single network drive.
    
.DESCRIPTION
    Maps a network drive with conflict resolution.
    Handles VPN scenarios (force remap), Retail restrictions.
    
.PARAMETER Mapping
    Drive mapping object from backend
    
.PARAMETER CurrentDrives
    Array of currently mapped drives
    
.OUTPUTS
    Boolean - true if successful
#>
Function Set-DMDriveMapping {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$Mapping,
        
        [Parameter(Mandatory=$False)]
        [Array]$CurrentDrives = @(),
        
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$UserInfo = $Null
    )
    
    Try {
        # Validate mapping has required properties
        If ($Null -eq $Mapping -or [String]::IsNullOrEmpty($Mapping.DriveLetter)) {
            Write-DMLog "Mapper Drive: Skipping mapping with empty or null DriveLetter" -Level Warning
            Return $False
        }
        
        [String]$DriveLetter = $Mapping.DriveLetter.TrimEnd(':')
        [String]$UncPath = If (-not [String]::IsNullOrEmpty($Mapping.UncPath)) { Expand-DMEnvironmentPath -Path $Mapping.UncPath } Else { "" }
        [String]$Description = If (-not [String]::IsNullOrEmpty($Mapping.Description)) { $Mapping.Description } Else { "" }
        
        Write-DMLog "Mapper Drive: Processing $DriveLetter -> $UncPath [$Description]" -Level Verbose
        
        # Check if Retail user trying to map V: drive (restricted)
        [String]$UserDN = ""
        If ($Null -ne $UserInfo -and -not [String]::IsNullOrEmpty($UserInfo.DistinguishedName)) {
            $UserDN = $UserInfo.DistinguishedName
        }
        
        If (Test-DMRetailUser -DistinguishedName $UserDN) {
            If ($DriveLetter.ToUpper() -eq "V") {
                Write-DMLog "Mapper Drive: Skipping V drive mapping for Retail users" -Level Info
                Return $True
            }
        }
        
        # Check if drive is already mapped
        [Object]$ExistingDrive = $CurrentDrives | Where-Object { $_.DriveLetter.TrimEnd(':') -eq $DriveLetter }
        
        If ($Null -ne $ExistingDrive) {
            # Drive letter is already in use
            If ($ExistingDrive.UncPath.ToUpper() -eq $UncPath.ToUpper()) {
                # Already mapped to correct path
                [Boolean]$IsVPN = Test-DMVPNConnection
                
                If ($IsVPN) {
                    Write-DMLog "Mapper Drive: Identified as VPN connected" -Level Info
                    Write-DMLog "Mapper Drive: About to remap '$DriveLetter' to '$UncPath'" -Level Info
                    
                    # Remove and remap
                    Remove-DMDriveMapping -DriveLetter $DriveLetter
                    Return Add-DMDriveMapping -DriveLetter $DriveLetter -UncPath $UncPath -Description $Description
                } Else {
                    Write-DMLog "Mapper Drive: '$DriveLetter' is already mapped to '$UncPath'. Will not take any action" -Level Info
                    Return $True
                }
            } Else {
                # Mapped to wrong path - remove and remap
                Write-DMLog "Mapper Drive: '$DriveLetter' is incorrectly mapped to '$($ExistingDrive.UncPath)' and will be removed" -Level Info
                Remove-DMDriveMapping -DriveLetter $DriveLetter
            }
        }
        
        # Map the drive
        Return Add-DMDriveMapping -DriveLetter $DriveLetter -UncPath $UncPath -Description $Description
    }
    Catch {
        [String]$DriveLetterForLog = If ($Null -ne $Mapping -and -not [String]::IsNullOrEmpty($Mapping.DriveLetter)) { $Mapping.DriveLetter } Else { "Unknown" }
        [String]$ErrorMessage = ""
        
        Try {
            If ($Null -ne $_.Exception -and -not [String]::IsNullOrEmpty($_.Exception.Message)) {
                $ErrorMessage = $_.Exception.Message
            } ElseIf ($Null -ne $_) {
                $ErrorMessage = $_.ToString()
            } Else {
                $ErrorMessage = "Unknown error occurred"
            }
        }
        Catch {
            $ErrorMessage = "Error details unavailable"
        }
        
        Write-DMLog "Mapper Drive: Error mapping $DriveLetterForLog - $ErrorMessage" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Adds a network drive mapping.
    
.DESCRIPTION
    Maps a network drive using WScript.Network and sets the description.
    
.PARAMETER DriveLetter
    Drive letter (without colon)
    
.PARAMETER UncPath
    UNC path to map
    
.PARAMETER Description
    Drive description/label
    
.OUTPUTS
    Boolean - true if successful
#>
Function Add-DMDriveMapping {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DriveLetter,
        
        [Parameter(Mandatory=$True)]
        [String]$UncPath,
        
        [Parameter(Mandatory=$False)]
        [String]$Description = ""
    )
    
    Try {
        # Validate parameters
        If ([String]::IsNullOrEmpty($DriveLetter)) {
            Write-DMLog "Mapper Drive: Cannot map drive - DriveLetter is empty" -Level Error
            Return $False
        }
        
        If ([String]::IsNullOrEmpty($UncPath)) {
            Write-DMLog "Mapper Drive: Cannot map drive '$DriveLetter' - UncPath is empty" -Level Error
            Return $False
        }
        
        Write-DMLog "Mapper Drive: About to map '$DriveLetter' to '$UncPath'" -Level Verbose
        
        # Use WScript.Network for compatibility
        [Object]$Network = New-Object -ComObject WScript.Network
        $Network.MapNetworkDrive("${DriveLetter}:", $UncPath, $True)
        
        Write-DMLog "Mapper Drive: Mapped '$DriveLetter' to '$UncPath'" -Level Info
        
        # Set description if provided
        If (-not [String]::IsNullOrEmpty($Description)) {
            Try {
                Write-DMLog "Mapper Drive: About to set description for '$DriveLetter' to '$Description'" -Level Verbose
                
                [Object]$Shell = New-Object -ComObject Shell.Application
                $Shell.NameSpace("${DriveLetter}:").Self.Name = $Description
                
                Write-DMLog "Mapper Drive: Description for '$DriveLetter' set to '$Description'" -Level Info
            } Catch {
                Write-DMLog "Mapper Drive: Failed to set description for '$DriveLetter': $($_.Exception.Message)" -Level Warning
            }
        }
        
        Return $True
    }
    Catch {
        Write-DMLog "Mapper Drive: Failed to map '$DriveLetter' to '$UncPath': $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Removes a network drive mapping.
    
.DESCRIPTION
    Disconnects a mapped network drive.
    
.PARAMETER DriveLetter
    Drive letter to disconnect (without colon)
    
.OUTPUTS
    Boolean - true if successful
#>
Function Remove-DMDriveMapping {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DriveLetter
    )
    
    Try {
        Write-DMLog "Mapper Drive: About to remove mapping for '$DriveLetter'" -Level Verbose
        
        [Object]$Network = New-Object -ComObject WScript.Network
        $Network.RemoveNetworkDrive("${DriveLetter}:", $True, $True)
        
        Write-DMLog "Mapper Drive: Removed incorrectly mapped drive '$DriveLetter'" -Level Info
        Return $True
    }
    Catch {
        Write-DMLog "Mapper Drive: Failed to remove drive mapping for '$DriveLetter': $($_.Exception.Message)" -Level Warning
        Return $False
    }
}

<#
.SYNOPSIS
    Removes drive mappings matching a wildcard pattern.
    
.DESCRIPTION
    Disconnects drives whose UNC paths match the specified pattern.
    Used for DisconnectOnLogin functionality.
    
.PARAMETER Pattern
    Wildcard pattern to match (supports * and ?)
    
.OUTPUTS
    Boolean - true if at least one drive was disconnected
#>
Function Remove-DMDriveMappingByPattern {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Pattern
    )
    
    Try {
        Write-DMLog "Mapper Drive: Processing disconnect pattern: $Pattern" -Level Verbose
        
        [Array]$CurrentDrives = Get-DMMappedDrives
        [Int]$DisconnectCount = 0
        
        ForEach ($Drive in $CurrentDrives) {
            If (Test-DMWildcardMatch -Text $Drive.UncPath -Pattern $Pattern) {
                Write-DMLog "Mapper Drive: Drive $($Drive.DriveLetter) matches disconnect pattern, removing" -Level Info
                
                If (Remove-DMDriveMapping -DriveLetter $Drive.DriveLetter.TrimEnd(':')) {
                    $DisconnectCount++
                }
            }
        }
        
        If ($DisconnectCount -gt 0) {
            Write-DMLog "Mapper Drive: Disconnected $DisconnectCount drive(s) matching pattern" -Level Info
        }
        
        Return ($DisconnectCount -gt 0)
    }
    Catch {
        Write-DMLog "Mapper Drive: Error processing disconnect pattern: $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Maps the home drive (H:) to My Documents.
    
.DESCRIPTION
    Maps H: drive to the user's My Documents folder.
    Only enabled for specific regions (EU, MUM, AMERICAS).
    
.PARAMETER CurrentDrives
    Array of currently mapped drives
    
.OUTPUTS
    Boolean - true if mapped or not applicable
#>
Function Set-DMHomeDriveMapping {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [Array]$CurrentDrives = @()
    )
    
    Try {
        # Check if H: drive mapping is enabled for this region
        # TODO: Read from config - for now, implement logic
        # Enabled regions: EU, MUM, AMERICAS (from PROJECT_HISTORY.md)
        
        Write-DMLog "Mapper Drive: Checking home drive (H:) mapping eligibility" -Level Verbose
        
        # Get My Documents path from registry
        [String]$MyDocumentsPath = Get-DMRegistryValue -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" -Name "Personal" -DefaultValue ""
        
        If ([String]::IsNullOrEmpty($MyDocumentsPath)) {
            Write-DMLog "Mapper Drive: Could not retrieve My Documents location for home drive mapping" -Level Verbose
            Return $False
        }
        
        Write-DMLog "Mapper Drive: About to map home drive to '$MyDocumentsPath'" -Level Verbose
        
        # Map H: drive
        Return Set-DMDriveMapping -Mapping ([PSCustomObject]@{
            DriveLetter = "H:"
            UncPath = $MyDocumentsPath
            Description = "Home Drive"
            DisconnectOnLogin = $False
        }) -CurrentDrives $CurrentDrives
    }
    Catch {
        Write-DMLog "Mapper Drive: Error mapping home drive: $($_.Exception.Message)" -Level Warning
        Return $False
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Invoke-DMDriveMapper',
    'Set-DMDriveMapping',
    'Add-DMDriveMapping',
    'Remove-DMDriveMapping',
    'Remove-DMDriveMappingByPattern',
    'Set-DMHomeDriveMapping'
)

