<#
.SYNOPSIS
    Desktop Management Personal Folder (PST) Inventory Module
    
.DESCRIPTION
    Collects Outlook PST file locations and sends inventory to backend.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: InventoryPersonalFolders_W10.vbs, GatherMappings_W10.vbs (GetMappedPSTs)
#>

# Import required modules
Using Module ..\Services\DMInventoryService.psm1
Using Module ..\Framework\DMLogger.psm1
Using Module ..\Framework\DMRegistry.psm1
Using Module ..\Framework\DMCommon.psm1
Using Module ..\Framework\DMUser.psm1
Using Module ..\Utilities\Test-Environment.psm1

# Constants for Outlook registry parsing
$Script:HKEY_CURRENT_USER = 0x80000001
$Script:MASTER_CONFIG_KEY = "01023d0e"
$Script:MASTER_KEY_GUID = "9207f3e0a3b11019908b08002b2a56c2"
$Script:PST_CHECK_KEY = "00033009"
$Script:PST_LOCATION_KEY = "01023d00"
$Script:PST_FILENAME_KEY = "001f6700"
$Script:OUTLOOK_PROFILES_INDEX = "Software\Microsoft\Office\16.0\Outlook"
$Script:OUTLOOK_PROFILES_ROOT = "Software\Microsoft\Office\16.0\Outlook\Profiles"

<#
.SYNOPSIS
    Collects and sends PST file inventory to backend.
    
.DESCRIPTION
    Parses Outlook registry to find PST files and sends inventory to backend.
    
    Skip Conditions:
    - VPN is connected
    - User is a Retail user
    - User has no email address in LDAP
    - No Outlook profile exists
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Invoke-DMPersonalFolderInventory -UserInfo $User -ComputerInfo $Computer
#>
Function Invoke-DMPersonalFolderInventory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo
    )
    
    Try {
        Write-DMLog "Inventory Personal Folder: Starting" -Level Verbose
        
        # Check if VPN is connected
        [Boolean]$IsVPNConnected = Test-DMVPNConnection
        
        If ($IsVPNConnected) {
            Write-DMLog "Inventory Personal Folder: Identified as VPN connected" -Level Info
            Write-DMLog "Inventory Personal Folder: VPN connected Laptop is requested to skip this process" -Level Info
            Write-DMLog "Inventory Personal Folder: Completed" -Level Info
            Return $True
        }
        
        # Check if Retail user
        If (Test-DMRetailUser -DistinguishedName $UserInfo.DistinguishedName) {
            Write-DMLog "Inventory Personal Folder: User belongs to Retail OU, skipping script execution" -Level Info
            Write-DMLog "Inventory Personal Folder: Completed" -Level Info
            Return $True
        }
        
        Write-DMLog "Inventory Personal Folder: Identified as VPN not connected" -Level Verbose
        
        # Check if user has email address (skip if no email or not domain-joined)
        If ([String]::IsNullOrEmpty($UserInfo.DistinguishedName) -or [String]::IsNullOrEmpty($UserInfo.Domain)) {
            Write-DMLog "Inventory Personal Folder: User is not domain-joined, skipping" -Level Verbose
            Write-DMLog "Inventory Personal Folder: Completed" -Level Verbose
            Return $True
        }
        
        [String]$Email = Get-DMUserEmail -DistinguishedName $UserInfo.DistinguishedName -Domain $UserInfo.Domain
        
        If ([String]::IsNullOrEmpty($Email)) {
            Write-DMLog "Inventory Personal Folder: User has no email address in LDAP, skipping" -Level Verbose
            Write-DMLog "Inventory Personal Folder: Completed" -Level Verbose
            Return $True
        }
        
        Write-DMLog "Inventory Personal Folder: User email: $Email" -Level Verbose
        
        # Get mapped PST files
        [Array]$PSTFiles = Get-DMOutlookPSTFiles
        
        If ($PSTFiles.Count -eq 0) {
            Write-DMLog "Inventory Personal Folder: No PST files found" -Level Verbose
            Write-DMLog "Inventory Personal Folder: Completed" -Level Verbose
            Return $True
        }
        
        Write-DMLog "Inventory Personal Folder: Found $($PSTFiles.Count) PST file(s)" -Level Verbose
        
        # Send each PST to backend
        [Boolean]$OverallSuccess = $True
        
        ForEach ($PST in $PSTFiles) {
            Write-DMLog "Inventory Personal Folder: Sending: $($PST.Path)" -Level Verbose
            
            [Boolean]$Success = Send-DMPSTInventory -PSTInfo $PST -UserInfo $UserInfo -ComputerInfo $ComputerInfo
            
            If (-not $Success) {
                $OverallSuccess = $False
            }
        }
        
        If ($OverallSuccess) {
            Write-DMLog "Inventory Personal Folder: Successfully sent all PST inventory" -Level Verbose
        } Else {
            Write-DMLog "Inventory Personal Folder: Some PST inventory sends failed" -Level Warning
        }
        
        Write-DMLog "Inventory Personal Folder: Completed" -Level Verbose
        Return $OverallSuccess
    }
    Catch {
        Write-DMLog "Inventory Personal Folder: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Gets all Outlook PST files from registry.
    
.DESCRIPTION
    Parses Outlook registry to extract PST file paths.
    Supports Outlook 2013, 2016, 2019, 365.
    
.OUTPUTS
    Array of PST file objects
    
.EXAMPLE
    $PSTs = Get-DMOutlookPSTFiles
#>
Function Get-DMOutlookPSTFiles {
    [CmdletBinding()]
    Param()
    
    Try {
        [Array]$PSTList = @()
        
        # Get default Outlook profile
        [String]$DefaultProfile = Get-DMOutlookDefaultProfile
        
        If ([String]::IsNullOrEmpty($DefaultProfile)) {
            Write-DMLog "Get Outlook PSTs: No Outlook profile found" -Level Verbose
            Return @()
        }
        
        Write-DMLog "Get Outlook PSTs: Default profile: $DefaultProfile" -Level Verbose
        
        # Get StdRegProv WMI object for binary value reading
        [Object]$RegProv = Get-WmiObject -Namespace "root\default" -Class StdRegProv
        
        If ($Null -eq $RegProv) {
            Write-DMLog "Get Outlook PSTs: Could not get StdRegProv WMI object" -Level Verbose
            Return @()
        }
        
        # Read master config binary value to get PST GUIDs
        [String]$ProfilePath = "$Script:OUTLOOK_PROFILES_ROOT\$DefaultProfile\$Script:MASTER_KEY_GUID"
        [Object]$BinaryData = $Null
        
        Try {
            [Void]$RegProv.GetBinaryValue($Script:HKEY_CURRENT_USER, $ProfilePath, $Script:MASTER_CONFIG_KEY, [Ref]$BinaryData)
        } Catch {
            Write-DMLog "Get Outlook PSTs: Registry path does not exist or no PST data configured" -Level Verbose
            Return @()
        }
        
        If ($Null -eq $BinaryData -or $BinaryData.Count -eq 0) {
            Write-DMLog "Get Outlook PSTs: No master config data found" -Level Verbose
            Return @()
        }
        
        # Parse binary data to extract GUIDs
        [Array]$PSTGUIDs = Parse-DMOutlookGUIDs -BinaryData $BinaryData -ProfilePath "$Script:OUTLOOK_PROFILES_ROOT\$DefaultProfile"
        
        Write-DMLog "Get Outlook PSTs: Found $($PSTGUIDs.Count) potential PST GUID(s)" -Level Verbose
        
        # For each GUID, get PST file information
        ForEach ($GUID in $PSTGUIDs) {
            [String]$PSTPath = Get-DMPSTFilePath -GUID $GUID -ProfilePath "$Script:OUTLOOK_PROFILES_ROOT\$DefaultProfile" -RegProv $RegProv
            
            If ([String]::IsNullOrEmpty($PSTPath)) {
                Continue
            }
            
            # Get file information
            [Int64]$Size = Get-DMFileSize -Path $PSTPath
            [DateTime]$LastModified = Get-DMFileLastModified -Path $PSTPath
            [String]$LastModifiedString = ""
            
            If ($Null -ne $LastModified) {
                $LastModifiedString = $LastModified.ToString("yyyy-MM-ddTHH:mm:ss")
            }
            
            # Convert to UNC path if it's a mapped drive
            [String]$UncPath = Convert-DMPathToUNC -Path $PSTPath
            
            $PSTList += [PSCustomObject]@{
                PSTypeName = 'DM.PSTFile'
                Path = $PSTPath
                UncPath = $UncPath
                Size = $Size
                LastModified = $LastModifiedString
            }
            
            Write-DMLog "Get Outlook PSTs: Found PST: $PSTPath (Size: $Size bytes)" -Level Verbose
        }
        
        Return $PSTList
    }
    Catch {
        Write-DMLog "Get Outlook PSTs: Error - $($_.Exception.Message)" -Level Warning
        Return @()
    }
}

<#
.SYNOPSIS
    Gets the default Outlook profile name.
    
.DESCRIPTION
    Reads DefaultProfile value from Outlook registry.
    
.OUTPUTS
    String - profile name or empty string
#>
Function Get-DMOutlookDefaultProfile {
    [CmdletBinding()]
    Param()
    
    Try {
        [String]$ProfileName = Get-DMRegistryValue -Path "HKCU:\$Script:OUTLOOK_PROFILES_INDEX" -Name "DefaultProfile" -DefaultValue ""
        Return $ProfileName
    }
    Catch {
        Return ""
    }
}

<#
.SYNOPSIS
    Parses binary data to extract PST GUIDs.
    
.DESCRIPTION
    Converts binary data to hex and extracts 32-character GUIDs.
    
.PARAMETER BinaryData
    Binary data array
    
.PARAMETER ProfilePath
    Outlook profile registry path
    
.OUTPUTS
    Array of GUID strings
#>
Function Parse-DMOutlookGUIDs {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Byte[]]$BinaryData,
        
        [Parameter(Mandatory=$True)]
        [String]$ProfilePath
    )
    
    Try {
        [Array]$GUIDs = @()
        [String]$CurrentGUID = ""
        
        # Get StdRegProv for validation
        [Object]$RegProv = Get-WmiObject -Namespace "root\default" -Class StdRegProv
        
        ForEach ($Byte in $BinaryData) {
            # Convert byte to 2-character hex string
            [String]$HexValue = "{0:X2}" -f $Byte
            $CurrentGUID += $HexValue
            
            # When we have 32 hex characters, we have a complete GUID
            If ($CurrentGUID.Length -eq 32) {
                # Validate if this is a PST GUID
                If (Test-DMIsPSTGUID -GUID $CurrentGUID -ProfilePath $ProfilePath -RegProv $RegProv) {
                    $GUIDs += $CurrentGUID
                    Write-DMLog "Parse Outlook GUIDs: Found PST GUID: $CurrentGUID" -Level Verbose
                }
                
                # Reset for next GUID
                $CurrentGUID = ""
            }
        }
        
        Return $GUIDs
    }
    Catch {
        Write-DMLog "Parse Outlook GUIDs: Error - $($_.Exception.Message)" -Level Verbose
        Return @()
    }
}

<#
.SYNOPSIS
    Tests if a GUID represents a PST file.
    
.DESCRIPTION
    Checks the binary value at GUID\00033009 to validate it's a PST.
    
.PARAMETER GUID
    GUID to check
    
.PARAMETER ProfilePath
    Outlook profile path
    
.PARAMETER RegProv
    StdRegProv WMI object
    
.OUTPUTS
    Boolean - true if PST GUID
#>
Function Test-DMIsPSTGUID {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$GUID,
        
        [Parameter(Mandatory=$True)]
        [String]$ProfilePath,
        
        [Parameter(Mandatory=$True)]
        [Object]$RegProv
    )
    
    Try {
        [String]$GUIDPath = "$ProfilePath\$GUID"
        [Object]$CheckData = $Null
        
        [Void]$RegProv.GetBinaryValue($Script:HKEY_CURRENT_USER, $GUIDPath, $Script:PST_CHECK_KEY, [Ref]$CheckData)
        
        If ($Null -eq $CheckData) {
            Return $False
        }
        
        # Calculate hex length sum (VBScript logic: sum of all hex values)
        [Int]$HexSum = 0
        ForEach ($Byte in $CheckData) {
            $HexSum += [Convert]::ToInt32($Byte)
        }
        
        # VBScript checks if sum equals 20
        Return ($HexSum -eq 20)
    }
    Catch {
        Return $False
    }
}

<#
.SYNOPSIS
    Gets PST file path from GUID.
    
.DESCRIPTION
    Reads binary data from registry to extract PST file path.
    
.PARAMETER GUID
    PST GUID
    
.PARAMETER ProfilePath
    Outlook profile path
    
.PARAMETER RegProv
    StdRegProv WMI object
    
.OUTPUTS
    String - PST file path
#>
Function Get-DMPSTFilePath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$GUID,
        
        [Parameter(Mandatory=$True)]
        [String]$ProfilePath,
        
        [Parameter(Mandatory=$True)]
        [Object]$RegProv
    )
    
    Try {
        [String]$GUIDPath = "$ProfilePath\$GUID"
        
        # Get location GUID
        [Object]$LocationData = $Null
        [Void]$RegProv.GetBinaryValue($Script:HKEY_CURRENT_USER, $GUIDPath, $Script:PST_LOCATION_KEY, [Ref]$LocationData)
        
        If ($Null -eq $LocationData) {
            Return ""
        }
        
        # Convert location binary to hex string (GUID)
        [String]$LocationGUID = ""
        ForEach ($Byte in $LocationData) {
            $LocationGUID += "{0:X2}" -f $Byte
        }
        
        # Get file name from location GUID
        [String]$LocationPath = "$ProfilePath\$LocationGUID"
        [Object]$FileNameData = $Null
        [Void]$RegProv.GetBinaryValue($Script:HKEY_CURRENT_USER, $LocationPath, $Script:PST_FILENAME_KEY, [Ref]$FileNameData)
        
        If ($Null -eq $FileNameData) {
            Return ""
        }
        
        # Convert binary data to string (each byte is a character)
        [String]$FileName = ""
        ForEach ($Byte in $FileNameData) {
            If ($Byte -gt 0) {
                $FileName += [Char]$Byte
            }
        }
        
        Return $FileName.Trim()
    }
    Catch {
        Write-DMLog "Get PST File Path: Error - $($_.Exception.Message)" -Level Verbose
        Return ""
    }
}

<#
.SYNOPSIS
    Converts a file path to UNC path if on mapped drive.
    
.DESCRIPTION
    If the path is on a mapped network drive, converts to UNC path.
    
.PARAMETER Path
    File path to convert
    
.OUTPUTS
    String - UNC path or original path
    
.EXAMPLE
    $UncPath = Convert-DMPathToUNC -Path "H:\Data\Archive.pst"
#>
Function Convert-DMPathToUNC {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Path
    )
    
    Try {
        If ([String]::IsNullOrEmpty($Path) -or $Path.Length -lt 2) {
            Return $Path
        }
        
        # Check if path starts with a drive letter
        If ($Path[1] -ne ':') {
            # Not a drive letter path (might already be UNC)
            Return $Path
        }
        
        [String]$DriveLetter = $Path.Substring(0, 1)
        
        # Get UNC path for this drive
        [String]$UncPath = ConvertTo-DMUNCPath -DriveLetter $DriveLetter
        
        If ([String]::IsNullOrEmpty($UncPath)) {
            # Not a network drive
            Return $Path
        }
        
        # Replace drive letter with UNC path
        [String]$RelativePath = $Path.Substring(2)  # Skip "X:"
        [String]$FullUncPath = $UncPath + $RelativePath
        
        Return $FullUncPath
    }
    Catch {
        Return $Path
    }
}

<#
.SYNOPSIS
    Sends PST inventory to backend.
    
.DESCRIPTION
    Sends individual PST file information to the inventory service.
    
.PARAMETER PSTInfo
    PST file information object
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.OUTPUTS
    Boolean - true if successful
#>
Function Send-DMPSTInventory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$PSTInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo
    )
    
    Try {
        # Get inventory server
        [Object]$Server = Get-DMInventoryServer -Domain $ComputerInfo.Domain
        
        If ($Null -eq $Server -or -not $Server.ServiceAvailable) {
            Write-DMLog "Send PST Inventory: No available inventory service found" -Level Warning
            Return $False
        }
        
        # Build method body XML
        [String]$MethodBody = ""
        $MethodBody += New-DMXMLElement -ElementName "UserId" -Content $UserInfo.Name
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "HostName" -Content $ComputerInfo.Name
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "Path" -Content $PSTInfo.Path
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "UncPath" -Content $PSTInfo.UncPath
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "Size" -Content $PSTInfo.Size.ToString()
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "PstLastUpdate" -Content $PSTInfo.LastModified
        
        # Create SOAP envelope
        [String]$SoapRequest = New-DMSOAPEnvelope -MethodName "InsertActivePersonalFolderMappingsFromInventory" -MethodBody $MethodBody
        
        # Send request
        [Object]$Response = Invoke-DMSOAPRequest -ServiceURL $Server.ServiceURL -SoapRequest $SoapRequest -Timeout $Server.Timeout
        
        Return $Response.Success
    }
    Catch {
        Write-DMLog "Send PST Inventory: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Invoke-DMPersonalFolderInventory',
    'Get-DMOutlookPSTFiles',
    'Get-DMOutlookDefaultProfile',
    'Convert-DMPathToUNC',
    'Send-DMPSTInventory'
)

