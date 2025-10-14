<#
.SYNOPSIS
    Desktop Management Printer Mapper Module
    
.DESCRIPTION
    Maps network printers based on backend configuration.
    Handles cleanup of unmapped printers and default printer assignment.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: MapperPrinters_W10.vbs
#>

# Import required modules
Using Module ..\Services\DMMapperService.psm1
Using Module ..\Framework\DMLogger.psm1
Using Module ..\Utilities\Test-Environment.psm1
Using Module ..\Services\DMServiceCommon.psm1

<#
.SYNOPSIS
    Maps network printers based on backend configuration.
    
.DESCRIPTION
    Retrieves printer mappings from backend and applies them to the computer.
    
    Skip Conditions:
    - Laptop Offline PC group member
    - JAPAN domain user (unless in Regional Printer Mapping Inclusion group)
    - Regional Printer Mapping Exclusion group member
    - Server OS
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Invoke-DMPrinterMapper -UserInfo $User -ComputerInfo $Computer
#>
Function Invoke-DMPrinterMapper {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo
    )
    
    Try {
        Write-DMLog "Mapper Printer: Starting printer mapping process" -Level Info
        
        # Check if offline laptop (skip mapping)
        If (Test-DMHostPartOfGroup -ComputerInfo $ComputerInfo -GroupName "Laptop Offline PC") {
            Write-DMLog "Mapper Printer: Identified as an Offline Laptop" -Level Info
            Write-DMLog "Mapper Printer: Offline Laptop is requested to skip this process" -Level Info
            Write-DMLog "Mapper Printer: Completed" -Level Info
            Return $True
        }
        
        # Check if user is in inclusion group (overrides JAPAN domain restriction)
        [Boolean]$IsInInclusionGroup = Test-DMUserPartOfGroup -UserInfo $UserInfo -GroupName "Regional Printer Mapping Inclusion"
        
        If ($IsInInclusionGroup) {
            Write-DMLog "Mapper Printer: User is part of the regional printer mapping inclusion group, printer mapping will continue" -Level Info
        }
        
        # Check if JAPAN domain (disabled by default unless in inclusion group)
        [Array]$JapanDomains = @("JAPAN", "QAJAPAN", "RNDJAPAN")
        If ($JapanDomains -contains $UserInfo.ShortDomain.ToUpper() -and -not $IsInInclusionGroup) {
            Write-DMLog "Mapper Printer: User belongs to the domain '$($UserInfo.ShortDomain)', printer mapping functionality is disabled by default" -Level Info
            Write-DMLog "Mapper Printer: Completed" -Level Info
            Return $True
        }
        
        # Check if user is in exclusion group
        If (Test-DMUserPartOfGroup -UserInfo $UserInfo -GroupName "Regional Printer Mapping Exclusion") {
            Write-DMLog "Mapper Printer: User is part of the regional printer mapping exclusion group, skipping" -Level Info
            Write-DMLog "Mapper Printer: Completed" -Level Info
            Return $True
        }
        
        # Check if server OS (printers not mapped on servers)
        If (Test-DMServerOS) {
            Write-DMLog "Mapper Printer: Operating system is Server, mapper printer script will not run" -Level Info
            Return $True
        }
        
        # Get printer mappings from backend (uses COMPUTER groups, not user)
        [Array]$PrinterMappings = Get-DMPrinterMappings -ComputerInfo $ComputerInfo
        
        If ($PrinterMappings.Count -eq 0) {
            Write-DMLog "Mapper Printer: No printer mappings returned from service" -Level Info
            Return $True
        }
        
        Write-DMLog "Mapper Printer: Retrieved $($PrinterMappings.Count) printer mapping(s) from service" -Level Info
        
        # Get currently installed printers
        [Array]$CurrentPrinters = Get-DMInstalledPrinters
        
        [Int]$SuccessfulMappings = 0
        [Int]$FailedMappings = 0
        
        # Map each printer
        ForEach ($Mapping in $PrinterMappings) {
            [Boolean]$Success = Add-DMPrinterMapping -Mapping $Mapping -CurrentPrinters $CurrentPrinters
            If ($Success) { $SuccessfulMappings++ } Else { $FailedMappings++ }
        }
        
        # Cleanup printers not in backend configuration
        If ($CurrentPrinters.Count -gt 0) {
            [Int]$RemovedCount = Remove-DMUnmanagedPrinters -ManagedPrinters $PrinterMappings -CurrentPrinters $CurrentPrinters
        } Else {
            [Int]$RemovedCount = 0
        }
        
        Write-DMLog "Mapper Printer: Summary - Mapped: $SuccessfulMappings, Failed: $FailedMappings, Removed: $RemovedCount" -Level Info
        Write-DMLog "Mapper Printer: Completed" -Level Info
        
        Return ($FailedMappings -eq 0)
    }
    Catch {
        Write-DMLog "Mapper Printer: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Adds a network printer mapping.
    
.DESCRIPTION
    Connects to a network printer and optionally sets it as default.
    
.PARAMETER Mapping
    Printer mapping object from backend
    
.PARAMETER CurrentPrinters
    Array of currently installed printers
    
.OUTPUTS
    Boolean - true if successful
#>
Function Add-DMPrinterMapping {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$Mapping,
        
        [Parameter(Mandatory=$False)]
        [Array]$CurrentPrinters = @()
    )
    
    Try {
        [String]$UncPath = $Mapping.UncPath
        [Boolean]$IsDefault = $Mapping.IsDefault
        
        Write-DMLog "Mapper Printer: Processing printer: $UncPath" -Level Verbose
        
        # Check if printer is already installed
        [Object]$ExistingPrinter = $CurrentPrinters | Where-Object { $_.UncPath -eq $UncPath }
        
        If ($Null -ne $ExistingPrinter) {
            Write-DMLog "Mapper Printer: Printer '$UncPath' is already installed" -Level Verbose
            
            # Set as default if required
            If ($IsDefault) {
                Return Set-DMDefaultPrinter -UncPath $UncPath
            }
            
            Return $True
        }
        
        # Add printer connection
        Write-DMLog "Mapper Printer: About to add printer: $UncPath" -Level Verbose
        
        [Object]$Network = New-Object -ComObject WScript.Network
        $Network.AddWindowsPrinterConnection($UncPath)
        
        Write-DMLog "Mapper Printer: Added printer: $UncPath" -Level Info
        
        # Set as default if required
        If ($IsDefault) {
            Set-DMDefaultPrinter -UncPath $UncPath
        }
        
        Return $True
    }
    Catch {
        Write-DMLog "Mapper Printer: Failed to add printer '$($Mapping.UncPath)': $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Sets a printer as the default printer.
    
.DESCRIPTION
    Sets the specified network printer as the default printer.
    
.PARAMETER UncPath
    Printer UNC path
    
.OUTPUTS
    Boolean - true if successful
#>
Function Set-DMDefaultPrinter {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$UncPath
    )
    
    Try {
        Write-DMLog "Mapper Printer: Setting default printer to: $UncPath" -Level Verbose
        
        [Object]$Network = New-Object -ComObject WScript.Network
        $Network.SetDefaultPrinter($UncPath)
        
        Write-DMLog "Mapper Printer: Set default printer to: $UncPath" -Level Info
        Return $True
    }
    Catch {
        Write-DMLog "Mapper Printer: Failed to set default printer: $($_.Exception.Message)" -Level Warning
        Return $False
    }
}

<#
.SYNOPSIS
    Removes a network printer connection.
    
.DESCRIPTION
    Disconnects a network printer.
    
.PARAMETER UncPath
    Printer UNC path to remove
    
.OUTPUTS
    Boolean - true if successful
#>
Function Remove-DMPrinterMapping {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$UncPath
    )
    
    Try {
        Write-DMLog "Mapper Printer: Removing printer: $UncPath" -Level Verbose
        
        [Object]$Network = New-Object -ComObject WScript.Network
        $Network.RemovePrinterConnection($UncPath, $True, $True)
        
        Write-DMLog "Mapper Printer: Removed printer: $UncPath" -Level Info
        Return $True
    }
    Catch {
        Write-DMLog "Mapper Printer: Failed to remove printer '$UncPath': $($_.Exception.Message)" -Level Warning
        Return $False
    }
}

<#
.SYNOPSIS
    Gets all currently installed network printers.
    
.DESCRIPTION
    Queries for all network printer connections.
    
.OUTPUTS
    Array of printer objects
#>
Function Get-DMInstalledPrinters {
    [CmdletBinding()]
    Param()
    
    Try {
        [Array]$PrinterList = @()
        
        # Use WScript.Network to enumerate printers
        [Object]$Network = New-Object -ComObject WScript.Network
        [Object]$PrinterConnections = $Network.EnumPrinterConnections()
        
        # PrinterConnections is a collection: [0]=Port, [1]=Printer, [2]=Port, [3]=Printer, ...
        For ([Int]$i = 0; $i -lt $PrinterConnections.Count; $i += 2) {
            [String]$UncPath = $PrinterConnections.Item($i + 1)
            
            If ($UncPath -like "\\*") {
                $PrinterList += [PSCustomObject]@{
                    PSTypeName = 'DM.InstalledPrinter'
                    UncPath = $UncPath
                }
            }
        }
        
        Return $PrinterList
    }
    Catch {
        Write-DMLog "Get Installed Printers: Error - $($_.Exception.Message)" -Level Verbose
        Return @()
    }
}

<#
.SYNOPSIS
    Removes printers that are not in the backend configuration.
    
.DESCRIPTION
    Compares locally installed printers with backend configuration.
    Removes printers that are not managed by the backend.
    
.PARAMETER ManagedPrinters
    Array of printer mappings from backend
    
.PARAMETER CurrentPrinters
    Array of currently installed printers
    
.OUTPUTS
    Int - count of printers removed
#>
Function Remove-DMUnmanagedPrinters {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Array]$ManagedPrinters,
        
        [Parameter(Mandatory=$True)]
        [Array]$CurrentPrinters
    )
    
    Try {
        [Int]$RemovedCount = 0
        
        # Build list of managed printer UNC paths
        [Array]$ManagedPaths = $ManagedPrinters | ForEach-Object { $_.UncPath.ToUpper() }
        
        # Check each installed printer
        ForEach ($Printer in $CurrentPrinters) {
            [String]$UncPath = $Printer.UncPath.ToUpper()
            
            If ($ManagedPaths -notcontains $UncPath) {
                Write-DMLog "Mapper Printer: Printer '$($Printer.UncPath)' is not managed, removing" -Level Info
                
                If (Remove-DMPrinterMapping -UncPath $Printer.UncPath) {
                    $RemovedCount++
                }
            }
        }
        
        Return $RemovedCount
    }
    Catch {
        Write-DMLog "Remove Unmanaged Printers: Error - $($_.Exception.Message)" -Level Warning
        Return 0
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Invoke-DMPrinterMapper',
    'Add-DMPrinterMapping',
    'Set-DMDefaultPrinter',
    'Remove-DMPrinterMapping',
    'Get-DMInstalledPrinters',
    'Remove-DMUnmanagedPrinters'
)

