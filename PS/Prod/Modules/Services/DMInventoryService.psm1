<#
.SYNOPSIS
    Desktop Management Inventory Service Module
    
.DESCRIPTION
    Handles sending inventory data TO the backend service.
    Supports logon/logoff tracking, drive inventory, printer inventory, PST inventory.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: Inventory*_W10.vbs modules
#>

# Import required modules
Using Module .\DMServiceCommon.psm1
Using Module ..\Framework\DMLogger.psm1

<#
.SYNOPSIS
    Sends user session logon inventory to backend.
    
.DESCRIPTION
    Records user logon event with computer and user information.
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.PARAMETER Server
    Optional server object (if not provided, will discover automatically)
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Send-DMLogonInventory -UserInfo $User -ComputerInfo $Computer
#>
Function Send-DMLogonInventory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo,
        
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$Server = $Null
    )
    
    Try {
        Write-DMLog "Inventory Logon: Preparing to insert user session data" -Level Verbose
        
        # Get server if not provided
        If ($Null -eq $Server) {
            $Server = Get-DMServiceServer -ServiceName "ClassicInventory.asmx" -Domain $ComputerInfo.Domain
        }
        
        If ($Null -eq $Server -or -not $Server.ServiceAvailable) {
            Write-DMLog "Inventory Logon: No available inventory service found" -Level Warning
            Return $False
        }
        
        Write-DMLog "Inventory Logon: Using service: $($Server.ServiceURL)" -Level Verbose
        
        # Build SOAP request matching VBScript format (no headers, simple envelope)
        [String]$SOAPBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Body>
        <InsertLogonInventory xmlns="http://webtools.japan.nom">
            <Sessions>
                <InventoryUserSession>
                    <UserId>$($UserInfo.Name)</UserId>
                    <UserDomain>$($UserInfo.Domain)</UserDomain>
                    <HostName>$($ComputerInfo.Name)</HostName>
                    <Domain>$($ComputerInfo.Domain)</Domain>
                    <SiteName>$($ComputerInfo.Site)</SiteName>
                    <City>$($ComputerInfo.CityCode)</City>
                    <OuMapping>$($UserInfo.OUMapping)</OuMapping>
                </InventoryUserSession>
            </Sessions>
        </InsertLogonInventory>
    </soap:Body>
</soap:Envelope>
"@
        
        Write-DMLog "Inventory Logon: Sending data - User: $($UserInfo.Name), Computer: $($ComputerInfo.Name)" -Level Verbose
        
        # Send request using proven authentication pattern
        [Object]$Response = Send-DMSOAPRequest -ServerUrl $Server.ServiceURL -SOAPBody $SOAPBody -SOAPAction "http://webtools.japan.nom/InsertLogonInventory" -Timeout $Server.Timeout
        
        If ($Response.Success) {
            Write-DMLog "Inventory Logon: Successfully sent session data" -Level Verbose
            Return $True
        } Else {
            Write-DMLog "Inventory Logon: Failed to send session data (Status: $($Response.StatusCode))" -Level Warning
            Return $False
        }
    }
    Catch {
        Write-DMLog "Inventory Logon: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Sends user session logoff inventory to backend.
    
.DESCRIPTION
    Records user logoff event.
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.PARAMETER Server
    Optional server object
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Send-DMLogoffInventory -UserInfo $User -ComputerInfo $Computer
#>
Function Send-DMLogoffInventory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo,
        
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$Server = $Null
    )
    
    Try {
        Write-DMLog "Inventory Logoff: Preparing to insert user session data" -Level Verbose
        
        # Get server if not provided
        If ($Null -eq $Server) {
            $Server = Get-DMServiceServer -ServiceName "ClassicInventory.asmx" -Domain $ComputerInfo.Domain
        }
        
        If ($Null -eq $Server -or -not $Server.ServiceAvailable) {
            Write-DMLog "Inventory Logoff: No available inventory service found" -Level Warning
            Return $False
        }
        
        # Build SOAP request matching VBScript format (no headers, simple envelope)
        [String]$SOAPBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Body>
        <InsertLogoffInventory xmlns="http://webtools.japan.nom">
            <Sessions>
                <InventoryUserSession>
                    <UserId>$($UserInfo.Name)</UserId>
                    <UserDomain>$($UserInfo.Domain)</UserDomain>
                    <HostName>$($ComputerInfo.Name)</HostName>
                    <Domain>$($ComputerInfo.Domain)</Domain>
                    <SiteName>$($ComputerInfo.Site)</SiteName>
                    <City>$($ComputerInfo.CityCode)</City>
                    <OuMapping>$($UserInfo.OUMapping)</OuMapping>
                </InventoryUserSession>
            </Sessions>
        </InsertLogoffInventory>
    </soap:Body>
</soap:Envelope>
"@
        
        Write-DMLog "Inventory Logoff: Sending data - User: $($UserInfo.Name), Computer: $($ComputerInfo.Name)" -Level Verbose
        
        # Send request using proven authentication pattern
        [Object]$Response = Send-DMSOAPRequest -ServerUrl $Server.ServiceURL -SOAPBody $SOAPBody -SOAPAction "http://webtools.japan.nom/InsertLogoffInventory" -Timeout $Server.Timeout
        
        If ($Response.Success) {
            Write-DMLog "Inventory Logoff: Successfully sent session data" -Level Verbose
            Return $True
        } Else {
            Write-DMLog "Inventory Logoff: Failed to send session data (Status: $($Response.StatusCode))" -Level Warning
            Return $False
        }
    }
    Catch {
        Write-DMLog "Inventory Logoff: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Sends drive mapping inventory to backend.
    
.DESCRIPTION
    Records current drive mappings for the user.
    
.PARAMETER DriveInfo
    Drive mapping information (array of drive objects)
    
.PARAMETER UserInfo
    User information
    
.PARAMETER ComputerInfo
    Computer information
    
.PARAMETER Server
    Optional server object
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Send-DMDriveInventory -DriveInfo $Drives -UserInfo $User -ComputerInfo $Computer
#>
Function Send-DMDriveInventory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Array]$DriveInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo,
        
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$Server = $Null
    )
    
    Try {
        If ($DriveInfo.Count -eq 0) {
            Write-DMLog "Drive Inventory: No drives to send" -Level Verbose
            Return $True
        }
        
        # Get server if not provided
        If ($Null -eq $Server) {
            $Server = Get-DMServiceServer -ServiceName "ClassicInventory.asmx" -Domain $ComputerInfo.Domain
        }
        
        If ($Null -eq $Server -or -not $Server.ServiceAvailable) {
            Write-DMLog "Drive Inventory: No available inventory service found" -Level Warning
            Return $False
        }
        
        # Build SOAP request matching VBScript format (no headers, simple envelope)
        # VBScript wraps all drives in <Mappings> tags
        [String]$MappingsXML = "<Mappings>`n"
        ForEach ($Drive in $DriveInfo) {
            $MappingsXML += "`t<InventoryDrive>`n"
            $MappingsXML += "`t`t<UserId>$($UserInfo.Name)</UserId>`n"
            $MappingsXML += "`t`t<HostName>$($ComputerInfo.Name)</HostName>`n"
            $MappingsXML += "`t`t<Domain>$($ComputerInfo.Domain)</Domain>`n"
            $MappingsXML += "`t`t<SiteName>$($ComputerInfo.Site)</SiteName>`n"
            $MappingsXML += "`t`t<City>$($ComputerInfo.CityCode)</City>`n"
            $MappingsXML += "`t`t<Drive>$($Drive.DriveLetter)</Drive>`n"
            $MappingsXML += "`t`t<UncPath>$($Drive.UncPath)</UncPath>`n"
            $MappingsXML += "`t`t<Description>$($Drive.Description)</Description>`n"
            $MappingsXML += "`t`t<OuMapping>$($UserInfo.OUMapping)</OuMapping>`n"
            $MappingsXML += "`t</InventoryDrive>`n"
        }
        $MappingsXML += "</Mappings>"
        
        [String]$SOAPBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Body>
        <InsertActiveDriveMappingsFromInventory xmlns="http://webtools.japan.nom">
            $MappingsXML
        </InsertActiveDriveMappingsFromInventory>
    </soap:Body>
</soap:Envelope>
"@
        
        [Object]$Response = Send-DMSOAPRequest -ServerUrl $Server.ServiceURL -SOAPBody $SOAPBody -SOAPAction "http://webtools.japan.nom/InsertActiveDriveMappingsFromInventory" -Timeout $Server.Timeout
        
        If ($Response.Success) {
            Write-DMLog "Drive Inventory: Successfully sent drive inventory" -Level Verbose
            Return $True
        } Else {
            Write-DMLog "Drive Inventory: Failed to send drive inventory (Status: $($Response.StatusCode))" -Level Warning
            Return $False
        }
    }
    Catch {
        Write-DMLog "Drive Inventory: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Gets inventory service server.
    
.DESCRIPTION
    Discovers and returns the inventory service server.
    
.PARAMETER Domain
    Domain name
    
.OUTPUTS
    PSCustomObject - server information
    
.EXAMPLE
    $Server = Get-DMInventoryServer -Domain $Computer.Domain
#>
Function Get-DMInventoryServer {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$Domain = ""
    )
    
    Return Get-DMServiceServer -ServiceName "ClassicInventory.asmx" -Domain $Domain
}

# Export module members
Export-ModuleMember -Function @(
    'Send-DMLogonInventory',
    'Send-DMLogoffInventory',
    'Send-DMDriveInventory',
    'Get-DMInventoryServer'
)

