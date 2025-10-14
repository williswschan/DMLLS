<#
.SYNOPSIS
    Desktop Management Personal Folder (PST) Mapper Module
    
.DESCRIPTION
    Maps Outlook PST files based on backend configuration.
    Uses Outlook COM automation to add/remove PST data stores.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: MapperPersonalFolders_W10.vbs
#>

# Import required modules
Using Module ..\Services\DMMapperService.psm1
Using Module ..\Framework\DMLogger.psm1
Using Module ..\Framework\DMCommon.psm1
Using Module ..\Utilities\Test-Environment.psm1
Using Module ..\Services\DMServiceCommon.psm1

<#
.SYNOPSIS
    Maps Outlook PST files based on backend configuration.
    
.DESCRIPTION
    Retrieves PST mappings from backend and adds them to Outlook.
    
    Skip Conditions:
    - Laptop Offline PC group member
    - VPN connected
    - Retail user
    - Server OS
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Invoke-DMPersonalFolderMapper -UserInfo $User -ComputerInfo $Computer
#>
Function Invoke-DMPersonalFolderMapper {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo
    )
    
    Try {
        Write-DMLog "Mapper PST: Starting PST mapping process" -Level Info
        
        # Check if offline laptop or VPN connected
        [Boolean]$IsOfflineLaptop = Test-DMUserPartOfGroup -UserInfo $UserInfo -GroupName "Laptop Offline PC"
        [Boolean]$IsVPN = Test-DMVPNConnection
        
        If ($IsOfflineLaptop -or $IsVPN) {
            Write-DMLog "Mapper PST: Identified as VPN connected or Offline Laptop" -Level Info
            Write-DMLog "Mapper PST: Request to skip this process" -Level Info
            Write-DMLog "Mapper PST: Skipping Completed" -Level Info
            Return $True
        }
        
        # Check if Retail user
        If (Test-DMRetailUser -DistinguishedName $UserInfo.DistinguishedName) {
            Write-DMLog "Mapper PST: User belongs to Retail OU, skipping script execution" -Level Info
            Write-DMLog "Mapper PST: Skipping Completed" -Level Info
            Return $True
        }
        
        # Check if server OS
        If (Test-DMServerOS) {
            Write-DMLog "Mapper PST: Operating system is Server, mapper PST script will not run" -Level Info
            Return $True
        }
        
        # Get PST mappings from backend
        [Array]$PSTMappings = Get-DMPSTMappings -UserInfo $UserInfo
        
        If ($PSTMappings.Count -eq 0) {
            Write-DMLog "Mapper PST: No PST mappings returned from service" -Level Info
            Return $True
        }
        
        Write-DMLog "Mapper PST: Retrieved $($PSTMappings.Count) PST mapping(s) from service" -Level Info
        
        # Process disconnect patterns first
        ForEach ($Mapping in $PSTMappings) {
            If ($Mapping.DisconnectOnLogin) {
                Remove-DMPSTMappingByPattern -Pattern $Mapping.UncPath
            }
        }
        
        [Int]$SuccessfulMappings = 0
        [Int]$FailedMappings = 0
        
        # Map PST files
        ForEach ($Mapping in $PSTMappings) {
            If (-not $Mapping.DisconnectOnLogin) {
                [Boolean]$Success = Add-DMPSTMapping -UncPath $Mapping.UncPath
                If ($Success) { $SuccessfulMappings++ } Else { $FailedMappings++ }
            }
        }
        
        Write-DMLog "Mapper PST: Summary - Mapped: $SuccessfulMappings, Failed: $FailedMappings" -Level Info
        Write-DMLog "Mapper PST: Completed" -Level Info
        
        Return ($FailedMappings -eq 0)
    }
    Catch {
        Write-DMLog "Mapper PST: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Adds a PST file to Outlook.
    
.DESCRIPTION
    Uses Outlook COM automation to add a PST data store.
    
.PARAMETER UncPath
    UNC path to PST file
    
.OUTPUTS
    Boolean - true if successful
#>
Function Add-DMPSTMapping {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$UncPath
    )
    
    Try {
        Write-DMLog "Mapper PST: About to add PST: $UncPath" -Level Verbose
        
        # Check if PST file exists
        If (-not (Test-Path -Path $UncPath -PathType Leaf)) {
            Write-DMLog "Mapper PST: PST file not found: $UncPath" -Level Warning
            Return $False
        }
        
        # Create Outlook COM object
        [Object]$Outlook = New-DMCOMObject -ProgId "Outlook.Application"
        
        If ($Null -eq $Outlook) {
            Write-DMLog "Mapper PST: Could not create Outlook.Application COM object" -Level Warning
            Return $False
        }
        
        # Get MAPI namespace
        [Object]$Namespace = $Outlook.GetNameSpace("MAPI")
        
        # Check if PST is already added
        [Boolean]$AlreadyAdded = $False
        ForEach ($Store in $Namespace.Stores) {
            If ($Store.FilePath -eq $UncPath) {
                $AlreadyAdded = $True
                Break
            }
        }
        
        If ($AlreadyAdded) {
            Write-DMLog "Mapper PST: PST already added: $UncPath" -Level Verbose
            Return $True
        }
        
        # Add PST store
        $Namespace.AddStore($UncPath)
        
        Write-DMLog "Mapper PST: Added PST: $UncPath" -Level Info
        
        # Release COM objects
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
        
        Return $True
    }
    Catch {
        Write-DMLog "Mapper PST: Failed to add PST '$UncPath': $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Removes PST files matching a pattern.
    
.DESCRIPTION
    Removes PST files from Outlook whose paths match the specified pattern.
    
.PARAMETER Pattern
    Wildcard pattern to match
    
.OUTPUTS
    Boolean - true if at least one PST was removed
#>
Function Remove-DMPSTMappingByPattern {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Pattern
    )
    
    Try {
        Write-DMLog "Mapper PST: Processing disconnect pattern: $Pattern" -Level Verbose
        
        # Create Outlook COM object
        [Object]$Outlook = New-DMCOMObject -ProgId "Outlook.Application"
        
        If ($Null -eq $Outlook) {
            Write-DMLog "Mapper PST: Could not create Outlook.Application COM object" -Level Warning
            Return $False
        }
        
        # Get MAPI namespace
        [Object]$Namespace = $Outlook.GetNameSpace("MAPI")
        
        [Int]$RemovedCount = 0
        
        # Check each store
        ForEach ($Store in $Namespace.Stores) {
            [String]$FilePath = $Store.FilePath
            
            If (-not [String]::IsNullOrEmpty($FilePath) -and (Test-DMWildcardMatch -Text $FilePath -Pattern $Pattern)) {
                Write-DMLog "Mapper PST: PST matches disconnect pattern, removing: $FilePath" -Level Info
                
                Try {
                    $Namespace.RemoveStore($Store)
                    $RemovedCount++
                } Catch {
                    Write-DMLog "Mapper PST: Failed to remove PST: $($_.Exception.Message)" -Level Warning
                }
            }
        }
        
        # Release COM objects
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Namespace) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
        
        If ($RemovedCount -gt 0) {
            Write-DMLog "Mapper PST: Removed $RemovedCount PST file(s) matching pattern" -Level Info
        }
        
        Return ($RemovedCount -gt 0)
    }
    Catch {
        Write-DMLog "Mapper PST: Error processing disconnect pattern: $($_.Exception.Message)" -Level Error
        Return $False
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Invoke-DMPersonalFolderMapper',
    'Add-DMPSTMapping',
    'Remove-DMPSTMappingByPattern'
)

