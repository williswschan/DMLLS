<#
.SYNOPSIS
    Desktop Management Drive Inventory Module
    
.DESCRIPTION
    Collects currently mapped network drives and sends inventory to backend.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: InventoryDrives_W10.vbs, GatherMappings_W10.vbs (GetMappedDrives)
#>

# Import required modules
Using Module ..\Services\DMInventoryService.psm1
Using Module ..\Framework\DMLogger.psm1
Using Module ..\Framework\DMRegistry.psm1
Using Module ..\Utilities\Test-Environment.psm1

<#
.SYNOPSIS
    Collects and sends drive mapping inventory to backend.
    
.DESCRIPTION
    Reads all currently mapped network drives and sends the inventory to the backend server.
    Skips if VPN is connected (offline laptop scenario).
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Invoke-DMDriveInventory -UserInfo $User -ComputerInfo $Computer
#>
Function Invoke-DMDriveInventory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$UserInfo = $Null,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo
    )
    
    Try {
        Write-DMLog "Inventory Drive: Starting" -Level Verbose
        
        # Check if VPN is connected
        [Boolean]$IsVPNConnected = Test-DMVPNConnection
        
        If ($IsVPNConnected) {
            Write-DMLog "Inventory Drive: Identified as VPN connected" -Level Info
            Write-DMLog "Inventory Drive: VPN connected Laptop is requested to skip this process" -Level Info
            Write-DMLog "Inventory Drive: Completed" -Level Info
            Return $True  # Return true because skipping is expected behavior
        }
        
        Write-DMLog "Inventory Drive: Identified as VPN not connected" -Level Verbose
        
        # Get all mapped drives
        [Array]$MappedDrives = Get-DMMappedDrives -UserInfo $UserInfo
        
        If ($MappedDrives.Count -eq 0) {
            Write-DMLog "Inventory Drive: No mapped drives found" -Level Verbose
            Write-DMLog "Inventory Drive: Completed" -Level Verbose
            Return $True
        }
        
        Write-DMLog "Inventory Drive: Found $($MappedDrives.Count) mapped drive(s)" -Level Verbose
        
        # Send drive inventory to backend (create minimal user info if not provided)
        If ($Null -eq $UserInfo) {
            $UserInfo = [PSCustomObject]@{
                Name = $env:USERNAME
                Domain = $env:USERDOMAIN
            }
        }
        
        [Boolean]$Success = Send-DMDriveInventory -DriveInfo $MappedDrives -UserInfo $UserInfo -ComputerInfo $ComputerInfo
        
        If ($Success) {
            Write-DMLog "Inventory Drive: Successfully sent drive inventory" -Level Verbose
        } Else {
            Write-DMLog "Inventory Drive: Failed to send drive inventory" -Level Warning
        }
        
        Write-DMLog "Inventory Drive: Completed" -Level Verbose
        Return $Success
    }
    Catch {
        Write-DMLog "Inventory Drive: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Gets all currently mapped network drives.
    
.DESCRIPTION
    Reads drive mappings from registry (HKCU\Network\<DriveLetter>\RemotePath)
    and retrieves drive descriptions/labels.
    
.PARAMETER UserInfo
    Optional user information for Retail user check
    
.OUTPUTS
    Array of drive mapping objects
    
.EXAMPLE
    $Drives = Get-DMMappedDrives
    $Drives = Get-DMMappedDrives -UserInfo $User
#>
Function Get-DMMappedDrives {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$UserInfo = $Null
    )
    
    Try {
        [Array]$DriveList = @()
        
        # Check if Network registry key exists
        [String]$NetworkKeyPath = "HKCU:\Network"
        
        If (-not (Test-Path -Path $NetworkKeyPath)) {
            Write-DMLog "Get Mapped Drives: No Network registry key found" -Level Verbose
            Return @()
        }
        
        # Get all subkeys (drive letters)
        [Array]$DriveLetters = Get-DMRegistrySubKeys -Path $NetworkKeyPath
        
        If ($DriveLetters.Count -eq 0) {
            Write-DMLog "Get Mapped Drives: No drive mappings in registry" -Level Verbose
            Return @()
        }
        
        Write-DMLog "Get Mapped Drives: Found $($DriveLetters.Count) drive mapping(s) in registry" -Level Verbose
        
        ForEach ($DriveLetter in $DriveLetters) {
            [String]$DriveKeyPath = "$NetworkKeyPath\$DriveLetter"
            
            # Get UNC path (RemotePath value)
            [String]$UncPath = Get-DMRegistryValue -Path $DriveKeyPath -Name "RemotePath" -DefaultValue ""
            
            If ([String]::IsNullOrEmpty($UncPath)) {
                Write-DMLog "Get Mapped Drives: Drive $DriveLetter has no RemotePath, skipping" -Level Verbose
                Continue
            }
            
            # Get drive description/label
            [String]$Description = Get-DMDriveDescription -DriveLetter $DriveLetter -UncPath $UncPath
            
            # Add to list
            $DriveList += [PSCustomObject]@{
                PSTypeName = 'DM.MappedDrive'
                DriveLetter = "${DriveLetter}:"
                UncPath = $UncPath
                Description = $Description
            }
            
            Write-DMLog "Get Mapped Drives: $DriveLetter -> $UncPath [$Description]" -Level Verbose
        }
        
        # Check for Retail users - special V: drive handling
        If ($Null -ne $UserInfo -and -not [String]::IsNullOrEmpty($UserInfo.DistinguishedName) -and (Test-DMRetailUser -DistinguishedName $UserInfo.DistinguishedName)) {
            [String]$HomeDrive = $env:HOMEDRIVE
            If (-not [String]::IsNullOrEmpty($HomeDrive) -and $HomeDrive -eq "V:") {
                # Check if V: is not already in the list
                [Boolean]$VDriveExists = $False
                ForEach ($Drive in $DriveList) {
                    If ($Drive.DriveLetter -eq "V:") {
                        $VDriveExists = $True
                        Break
                    }
                }
                
                If (-not $VDriveExists) {
                    Write-DMLog "Get Mapped Drives: Retail user - adding V: drive from HOMEDRIVE" -Level Verbose
                    $DriveList += [PSCustomObject]@{
                        PSTypeName = 'DM.MappedDrive'
                        DriveLetter = "V:"
                        UncPath = ""  # V: drive for retail may not have UNC path
                        Description = "Retail Home Drive"
                    }
                }
            }
        }
        
        Return $DriveList
    }
    Catch {
        Write-DMLog "Get Mapped Drives: Error - $($_.Exception.Message)" -Level Warning
        Return @()
    }
}

<#
.SYNOPSIS
    Gets the description/label for a mapped drive.
    
.DESCRIPTION
    Attempts to read drive description from registry MountPoints2.
    
.PARAMETER DriveLetter
    Drive letter (without colon)
    
.PARAMETER UncPath
    UNC path of the drive
    
.OUTPUTS
    String - drive description or empty string
    
.EXAMPLE
    $Description = Get-DMDriveDescription -DriveLetter "H" -UncPath "\\server\share"
#>
Function Get-DMDriveDescription {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DriveLetter,
        
        [Parameter(Mandatory=$True)]
        [String]$UncPath
    )
    
    Try {
        # Try to read from MountPoints2 registry
        # Path: HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2\##UncPath\LabelFromReg
        
        [String]$MountPointsPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\MountPoints2"
        
        # Convert UNC path to registry key format (replace \ with #)
        [String]$UncPathKey = $UncPath.Replace("\", "#")
        [String]$FullPath = "$MountPointsPath\$UncPathKey"
        
        If (Test-Path -Path $FullPath) {
            [String]$Label = Get-DMRegistryValue -Path $FullPath -Name "LabelFromReg" -DefaultValue ""
            If (-not [String]::IsNullOrEmpty($Label)) {
                Return $Label
            }
        }
        
        # Fallback: try to get from drive letter itself using Shell.Application
        Try {
            [Object]$Shell = New-Object -ComObject Shell.Application
            [Object]$Folder = $Shell.NameSpace("${DriveLetter}:")
            
            If ($Null -ne $Folder) {
                [String]$Name = $Folder.Self.Name
                If (-not [String]::IsNullOrEmpty($Name) -and $Name -ne "${DriveLetter}:") {
                    Return $Name
                }
            }
        } Catch {
            Write-DMLog "Get Drive Description: Shell.Application failed for $DriveLetter" -Level Verbose
        }
        
        Return ""
    }
    Catch {
        Write-DMLog "Get Drive Description: Error for $DriveLetter - $($_.Exception.Message)" -Level Verbose
        Return ""
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Invoke-DMDriveInventory',
    'Get-DMMappedDrives',
    'Get-DMDriveDescription'
)

