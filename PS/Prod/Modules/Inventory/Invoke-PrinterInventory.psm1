<#
.SYNOPSIS
    Desktop Management Printer Inventory Module
    
.DESCRIPTION
    Collects currently mapped network printers and sends inventory to backend.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: InventoryPrinters_W10.vbs
#>

# Import required modules
Using Module ..\Services\DMInventoryService.psm1
Using Module ..\Framework\DMLogger.psm1
Using Module ..\Utilities\Test-Environment.psm1
Using Module ..\Services\DMServiceCommon.psm1

<#
.SYNOPSIS
    Collects and sends printer inventory to backend.
    
.DESCRIPTION
    Reads all currently mapped network printers and sends the inventory to the backend server.
    
    Skip Conditions:
    - VPN is connected
    - User is part of "Regional Printer Mapping Exclusion" group
    - User is a Retail user
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Invoke-DMPrinterInventory -UserInfo $User -ComputerInfo $Computer
#>
Function Invoke-DMPrinterInventory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo
    )
    
    Try {
        Write-DMLog "Inventory Printer: Starting" -Level Verbose
        
        # Check if VPN is connected
        [Boolean]$IsVPNConnected = Test-DMVPNConnection
        
        If ($IsVPNConnected) {
            Write-DMLog "Inventory Printer: Identified as VPN connected" -Level Info
            Write-DMLog "Inventory Printer: VPN connected Laptop is requested to skip this process" -Level Info
            Write-DMLog "Inventory Printer: Completed" -Level Info
            Return $True
        }
        
        Write-DMLog "Inventory Printer: Identified as VPN not connected" -Level Verbose
        
        # Check if user is in Regional Printer Mapping Exclusion group
        If (Test-DMUserPartOfGroup -UserInfo $UserInfo -GroupName "Regional Printer Mapping Exclusion") {
            Write-DMLog "Inventory Printer: User is part of 'Regional Printer Mapping Exclusion' group, as such printer inventory process will be skipped" -Level Info
            Write-DMLog "Inventory Printer: Completed" -Level Info
            Return $True
        }
        
        # Check if Retail user
        If (Test-DMRetailUser -DistinguishedName $UserInfo.DistinguishedName) {
            Write-DMLog "Inventory Printer: User belongs to Retail OU, skipping script execution" -Level Info
            Write-DMLog "Inventory Printer: Completed" -Level Info
            Return $True
        }
        
        # Get all mapped network printers
        [Array]$NetworkPrinters = Get-DMNetworkPrinters
        
        If ($NetworkPrinters.Count -eq 0) {
            Write-DMLog "Inventory Printer: No network printers found" -Level Verbose
            Write-DMLog "Inventory Printer: Completed" -Level Verbose
            Return $True
        }
        
        Write-DMLog "Inventory Printer: Found $($NetworkPrinters.Count) network printer(s)" -Level Verbose
        
        # Send each printer to backend
        [Boolean]$OverallSuccess = $True
        
        ForEach ($Printer in $NetworkPrinters) {
            Write-DMLog "Inventory Printer: Sending: $($Printer.UncPath)" -Level Verbose
            
            [Boolean]$Success = Send-DMPrinterInventory -PrinterInfo $Printer -UserInfo $UserInfo -ComputerInfo $ComputerInfo
            
            If (-not $Success) {
                $OverallSuccess = $False
            }
        }
        
        If ($OverallSuccess) {
            Write-DMLog "Inventory Printer: Successfully sent all printer inventory" -Level Verbose
        } Else {
            Write-DMLog "Inventory Printer: Some printer inventory sends failed" -Level Warning
        }
        
        Write-DMLog "Inventory Printer: Completed" -Level Verbose
        Return $OverallSuccess
    }
    Catch {
        Write-DMLog "Inventory Printer: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Gets all currently mapped network printers.
    
.DESCRIPTION
    Queries WMI for network printers (Win32_Printer WHERE Network = TRUE).
    Collects UNC path, default status, driver, port, description.
    
.OUTPUTS
    Array of network printer objects
    
.EXAMPLE
    $Printers = Get-DMNetworkPrinters
#>
Function Get-DMNetworkPrinters {
    [CmdletBinding()]
    Param()
    
    Try {
        [Array]$PrinterList = @()
        
        # Query WMI for network printers
        [Array]$Printers = Get-CimInstance -ClassName Win32_Printer -Filter "Network = TRUE" -ErrorAction SilentlyContinue
        
        If ($Null -eq $Printers -or $Printers.Count -eq 0) {
            Write-DMLog "Get Network Printers: No network printers found" -Level Verbose
            Return @()
        }
        
        Write-DMLog "Get Network Printers: Found $($Printers.Count) network printer(s)" -Level Verbose
        
        # Get default printer name
        [String]$DefaultPrinterName = ""
        Try {
            [Object]$WshNetwork = New-Object -ComObject WScript.Network
            $DefaultPrinterName = $WshNetwork.UserName
        } Catch {
            Write-DMLog "Get Network Printers: Could not determine default printer via WScript.Network" -Level Verbose
        }
        
        ForEach ($Printer in $Printers) {
            # Determine if this is the default printer
            [Boolean]$IsDefault = ($Printer.Default -eq $True) -or ($Printer.Name -eq $DefaultPrinterName)
            
            # Extract UNC path (ShareName for network printers)
            [String]$UncPath = ""
            If (-not [String]::IsNullOrEmpty($Printer.ShareName)) {
                $UncPath = $Printer.ShareName
            } ElseIf ($Printer.PortName -like "\\*") {
                $UncPath = $Printer.PortName
            } ElseIf (-not [String]::IsNullOrEmpty($Printer.Name) -and $Printer.Name -like "\\*") {
                $UncPath = $Printer.Name
            }
            
            # Skip if no UNC path found
            If ([String]::IsNullOrEmpty($UncPath)) {
                Write-DMLog "Get Network Printers: Printer '$($Printer.Name)' has no UNC path, skipping" -Level Verbose
                Continue
            }
            
            [String]$Driver = If ($Null -ne $Printer.DriverName) { $Printer.DriverName } Else { "" }
            [String]$Port = If ($Null -ne $Printer.PortName) { $Printer.PortName } Else { "" }
            [String]$Description = If ($Null -ne $Printer.Comment) { $Printer.Comment } Else { "" }
            
            $PrinterList += [PSCustomObject]@{
                PSTypeName = 'DM.NetworkPrinter'
                UncPath = $UncPath
                IsDefault = $IsDefault
                Driver = $Driver
                Port = $Port
                Description = $Description
            }
            
            Write-DMLog "Get Network Printers: $UncPath $(If ($IsDefault) {'[DEFAULT]'})" -Level Verbose
        }
        
        Return $PrinterList
    }
    Catch {
        Write-DMLog "Get Network Printers: Error - $($_.Exception.Message)" -Level Warning
        Return @()
    }
}

<#
.SYNOPSIS
    Sends printer inventory to backend.
    
.DESCRIPTION
    Sends individual printer information to the inventory service.
    This is a helper wrapper around the service layer.
    
.PARAMETER PrinterInfo
    Printer information object
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.OUTPUTS
    Boolean - true if successful
#>
Function Send-DMPrinterInventory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$PrinterInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo
    )
    
    Try {
        # Get inventory server
        [Object]$Server = Get-DMInventoryServer -Domain $ComputerInfo.Domain
        
        If ($Null -eq $Server -or -not $Server.ServiceAvailable) {
            Write-DMLog "Send Printer Inventory: No available inventory service found" -Level Warning
            Return $False
        }
        
        # Build method body XML
        [String]$MethodBody = ""
        $MethodBody += New-DMXMLElement -ElementName "UserId" -Content $UserInfo.Name
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "HostName" -Content $ComputerInfo.Name
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "Domain" -Content $ComputerInfo.Domain
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "UncPath" -Content $PrinterInfo.UncPath
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "IsDefault" -Content $PrinterInfo.IsDefault.ToString().ToLower()
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "Driver" -Content $PrinterInfo.Driver
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "Port" -Content $PrinterInfo.Port
        $MethodBody += "`n"
        $MethodBody += New-DMXMLElement -ElementName "Description" -Content $PrinterInfo.Description
        
        # Create SOAP envelope
        [String]$SoapRequest = New-DMSOAPEnvelope -MethodName "InsertMapperPrinterInventory" -MethodBody $MethodBody
        
        # Send request
        [Object]$Response = Invoke-DMSOAPRequest -ServiceURL $Server.ServiceURL -SoapRequest $SoapRequest -Timeout $Server.Timeout
        
        Return $Response.Success
    }
    Catch {
        Write-DMLog "Send Printer Inventory: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Invoke-DMPrinterInventory',
    'Get-DMNetworkPrinters',
    'Send-DMPrinterInventory'
)

