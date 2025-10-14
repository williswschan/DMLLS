<#
.SYNOPSIS
    Desktop Management Mapper Service Module
    
.DESCRIPTION
    Handles getting mapping data FROM the backend service.
    Supports drive mappings, printer mappings, and PST mappings.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: Mapper*_W10.vbs modules
#>

# Import required modules
Using Module .\DMServiceCommon.psm1
Using Module ..\Framework\DMLogger.psm1

<#
.SYNOPSIS
    Gets drive mappings from backend for a user.
    
.DESCRIPTION
    Retrieves drive mapping configuration from the mapper service.
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER ComputerInfo
    Computer information object
    
.PARAMETER Server
    Optional server object (if not provided, will discover automatically)
    
.OUTPUTS
    Array of drive mapping objects
    
.EXAMPLE
    $Drives = Get-DMDriveMappings -UserInfo $User -ComputerInfo $Computer
#>
Function Get-DMDriveMappings {
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
        Write-DMLog "Mapper Drive: Retrieving drive mappings from service" -Level Verbose
        
        # Get server if not provided
        If ($Null -eq $Server) {
            $Server = Get-DMServiceServer -ServiceName "ClassicMapper.asmx" -Domain $ComputerInfo.Domain
        }
        
        If ($Null -eq $Server -or -not $Server.ServiceAvailable) {
            Write-DMLog "Mapper Drive: No available mapper service found" -Level Warning
            Return @()
        }
        
        Write-DMLog "Mapper Drive: Using service: $($Server.ServiceURL)" -Level Verbose
        
        # Build AD groups string
        [String]$AdGroupsString = ""
        If ($Null -ne $UserInfo.Groups -and $UserInfo.Groups.Count -gt 0) {
            [Array]$GroupNames = $UserInfo.Groups | ForEach-Object { $_.GroupName }
            $AdGroupsString = $GroupNames -join ","
        }
        
        # Build SOAP request using proven pattern from working code
        [String]$SOAPBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body>
  <GetUserDrives xmlns="http://webtools.japan.nom">
   <UserId xsi:type="xsd:string">$($UserInfo.Name)</UserId>
   <Domain xsi:type="xsd:string">$($UserInfo.ShortDomain)</Domain>
   <OuMapping xsi:type="xsd:string">$($UserInfo.OUMapping)</OuMapping>
   <AdGroups xsi:type="xsd:string">$($UserGroupsString)</AdGroups>
   <Site xsi:type="xsd:string">$($ComputerInfo.CityCode)</Site>
  </GetUserDrives>
 </soap:Body>
</soap:Envelope>
"@
        
        Write-DMLog "Mapper Drive: Requesting mappings for user: $($UserInfo.Name)" -Level Verbose
        
        # Send request using proven authentication pattern
        [Object]$Response = Send-DMSOAPRequestWithAuth -ServerUrl $Server.ServiceURL -SOAPBody $SOAPBody -SOAPAction "http://webtools.japan.nom/GetUserDrives" -Username $UserInfo.Name -Password "placeholder" -Timeout $Server.Timeout
        
        If (-not $Response.Success) {
            Write-DMLog "Mapper Drive: Failed to get mappings (Status: $($Response.StatusCode))" -Level Warning
            Return @()
        }
        
               # Parse response
               [Array]$Drives = Parse-DMDriveMappingsResponse -ResponseXML $Response.ResponseXML
               
               Write-DMLog "Mapper Drive: Retrieved $($Drives.Count) drive mapping(s)" -Level Info
               
               # Log detailed information for each drive mapping (Verbose level)
               If ($Drives.Count -gt 0) {
                   Write-DMLog "Mapper Drive: Detailed mapping information:" -Level Verbose
                   For ($i = 0; $i -lt $Drives.Count; $i++) {
                       $Drive = $Drives[$i]
                       Write-DMLog "Mapper Drive: Mapping $($i + 1) of $($Drives.Count):" -Level Verbose
                       Write-DMLog "Mapper Drive:   Id: '$($Drive.Id)'" -Level Verbose
                       Write-DMLog "Mapper Drive:   Domain: '$($Drive.Domain)'" -Level Verbose
                       Write-DMLog "Mapper Drive:   UserId: '$($Drive.UserId)'" -Level Verbose
                       Write-DMLog "Mapper Drive:   AdGroup: '$($Drive.AdGroup)'" -Level Verbose
                       Write-DMLog "Mapper Drive:   Site: '$($Drive.Site)'" -Level Verbose
                       Write-DMLog "Mapper Drive:   DriveLetter: '$($Drive.DriveLetter)' (Type: $($Drive.DriveLetter.GetType().Name), Length: $($Drive.DriveLetter.Length))" -Level Verbose
                       Write-DMLog "Mapper Drive:   UncPath: '$($Drive.UncPath)'" -Level Verbose
                       Write-DMLog "Mapper Drive:   Description: '$($Drive.Description)'" -Level Verbose
                       Write-DMLog "Mapper Drive:   DisconnectOnLogin: $($Drive.DisconnectOnLogin)" -Level Verbose
                       
                       # Also write to console for immediate visibility
                       Write-Host "Drive Mapping $($i + 1): $($Drive.DriveLetter) -> $($Drive.UncPath)" -ForegroundColor Cyan
                   }
               } Else {
                   Write-DMLog "Mapper Drive: No drive mappings found in response" -Level Warning
               }
        
        Return $Drives
    }
    Catch {
        Write-DMLog "Mapper Drive: Error - $($_.Exception.Message)" -Level Error
        Return @()
    }
}

<#
.SYNOPSIS
    Gets printer mappings from backend for a computer.
    
.DESCRIPTION
    Retrieves printer mapping configuration from the mapper service.
    Note: Printer mappings are based on COMPUTER groups, not user groups.
    
.PARAMETER ComputerInfo
    Computer information object
    
.PARAMETER Server
    Optional server object
    
.OUTPUTS
    Array of printer mapping objects
    
.EXAMPLE
    $Printers = Get-DMPrinterMappings -ComputerInfo $Computer
#>
Function Get-DMPrinterMappings {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo,
        
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$Server = $Null
    )
    
    Try {
        Write-DMLog "Mapper Printer: Retrieving printer mappings from service" -Level Verbose
        
        # Get server if not provided
        If ($Null -eq $Server) {
            $Server = Get-DMServiceServer -ServiceName "ClassicMapper.asmx" -Domain $ComputerInfo.Domain
        }
        
        If ($Null -eq $Server -or -not $Server.ServiceAvailable) {
            Write-DMLog "Mapper Printer: No available mapper service found" -Level Warning
            Return @()
        }
        
        # Build AD groups string (COMPUTER groups, not user)
        [String]$AdGroupsString = ""
        If ($Null -ne $ComputerInfo.Groups -and $ComputerInfo.Groups.Count -gt 0) {
            [Array]$GroupNames = $ComputerInfo.Groups | ForEach-Object { $_.GroupName }
            $AdGroupsString = $GroupNames -join ","
        }
        
        # Build SOAP request using proven pattern from working code
        [String]$SOAPBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body>
  <GetUserPrinters xmlns="http://webtools.japan.nom">
   <UserId xsi:type="xsd:string">$($ComputerInfo.Name)</UserId>
   <Domain xsi:type="xsd:string">$($ComputerInfo.ShortDomain)</Domain>
   <OuMapping xsi:type="xsd:string">$($ComputerInfo.OUMapping)</OuMapping>
   <AdGroups xsi:type="xsd:string">$($AdGroupsString)</AdGroups>
   <Site xsi:type="xsd:string">$($ComputerInfo.CityCode)</Site>
   <HostName xsi:type="xsd:string">$($ComputerInfo.Name)</HostName>
  </GetUserPrinters>
 </soap:Body>
</soap:Envelope>
"@
        
        Write-DMLog "Mapper Printer: Requesting mappings for computer: $($ComputerInfo.Name)" -Level Verbose
        
        # Send request using proven authentication pattern
        [Object]$Response = Send-DMSOAPRequestWithAuth -ServerUrl $Server.ServiceURL -SOAPBody $SOAPBody -SOAPAction "http://webtools.japan.nom/GetUserPrinters" -Username $ComputerInfo.Name -Password "placeholder" -Timeout $Server.Timeout
        
        If (-not $Response.Success) {
            Write-DMLog "Mapper Printer: Failed to get mappings (Status: $($Response.StatusCode))" -Level Warning
            Return @()
        }
        
        # Parse response
        [Array]$Printers = Parse-DMPrinterMappingsResponse -ResponseXML $Response.ResponseXML
        
        Write-DMLog "Mapper Printer: Retrieved $($Printers.Count) printer mapping(s)" -Level Verbose
        
        Return $Printers
    }
    Catch {
        Write-DMLog "Mapper Printer: Error - $($_.Exception.Message)" -Level Error
        Return @()
    }
}

<#
.SYNOPSIS
    Gets PST mappings from backend for a user.
    
.DESCRIPTION
    Retrieves PST (personal folder) mapping configuration from the mapper service.
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER Server
    Optional server object
    
.OUTPUTS
    Array of PST mapping objects
    
.EXAMPLE
    $PSTs = Get-DMPSTMappings -UserInfo $User
#>
Function Get-DMPSTMappings {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$ComputerInfo = $Null,
        
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$Server = $Null
    )
    
    Try {
        Write-DMLog "Mapper PST: Retrieving PST mappings from service" -Level Verbose
        
        # Get server if not provided
        If ($Null -eq $Server) {
            $Server = Get-DMServiceServer -ServiceName "ClassicMapper.asmx" -Domain $UserInfo.Domain
        }
        
        If ($Null -eq $Server -or -not $Server.ServiceAvailable) {
            Write-DMLog "Mapper PST: No available mapper service found" -Level Warning
            Return @()
        }
        
        # Build AD groups string
        [String]$UserGroupsString = ""
        If ($Null -ne $UserInfo.Groups -and $UserInfo.Groups.Count -gt 0) {
            [Array]$GroupNames = $UserInfo.Groups | ForEach-Object { $_.GroupName }
            $UserGroupsString = $GroupNames -join ","
        }
        
        # Get computer info if not provided
        If ($Null -eq $ComputerInfo) {
            $ComputerInfo = Get-DMComputerInfo
        }
        
        # Build SOAP request using proven pattern from working code
        [String]$SOAPBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body>
  <GetUserPersonalFolders xmlns="http://webtools.japan.nom">
   <UserId xsi:type="xsd:string">$($UserInfo.Name)</UserId>
   <Domain xsi:type="xsd:string">$($UserInfo.ShortDomain)</Domain>
   <OuMapping xsi:type="xsd:string">$($UserInfo.OUMapping)</OuMapping>
   <AdGroups xsi:type="xsd:string">$($UserGroupsString)</AdGroups>
   <Site xsi:type="xsd:string">$($UserInfo.CityCode)</Site>
   <HostName xsi:type="xsd:string">$($ComputerInfo.Name)</HostName>
  </GetUserPersonalFolders>
 </soap:Body>
</soap:Envelope>
"@
        
        Write-DMLog "Mapper PST: Requesting mappings for user: $($UserInfo.Name)" -Level Verbose
        
        # Send request using proven authentication pattern
        [Object]$Response = Send-DMSOAPRequestWithAuth -ServerUrl $Server.ServiceURL -SOAPBody $SOAPBody -SOAPAction "http://webtools.japan.nom/GetUserPersonalFolders" -Username $UserInfo.Name -Password "placeholder" -Timeout $Server.Timeout
        
        If (-not $Response.Success) {
            Write-DMLog "Mapper PST: Failed to get mappings (Status: $($Response.StatusCode))" -Level Warning
            Return @()
        }
        
        # Parse response
        [Array]$PSTs = Parse-DMPSTMappingsResponse -ResponseXML $Response.ResponseXML
        
        Write-DMLog "Mapper PST: Retrieved $($PSTs.Count) PST mapping(s)" -Level Verbose
        
        Return $PSTs
    }
    Catch {
        Write-DMLog "Mapper PST: Error - $($_.Exception.Message)" -Level Error
        Return @()
    }
}

<#
.SYNOPSIS
    Parses drive mappings from SOAP response XML.
    
.DESCRIPTION
    Extracts drive mapping objects from the SOAP response.
    
.PARAMETER ResponseXML
    SOAP response XML
    
.OUTPUTS
    Array of drive mapping objects
#>
Function Parse-DMDriveMappingsResponse {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Xml]$ResponseXML
    )
    
    Try {
        [Array]$Drives = @()
        
        # Get result node
        [Object]$ResultNode = Get-DMSOAPResult -ResponseXML $ResponseXML -MethodName "GetUserDrives"
        
        If ($Null -eq $ResultNode) {
            Return @()
        }
        
        # Parse Drive elements - VBScript uses GetUserDrivesResult structure with MapperDrive elements
        [Array]$DriveNodes = $ResultNode.SelectNodes("./*[local-name()='MapperDrive']")
        
        ForEach ($DriveNode in $DriveNodes) {
            [Object]$IdNode = $DriveNode.SelectSingleNode("*[local-name()='Id']")
            [String]$Id = If ($Null -ne $IdNode) { $IdNode.InnerText } Else { "" }
            
            [Object]$DomainNode = $DriveNode.SelectSingleNode("*[local-name()='Domain']")
            [String]$Domain = If ($Null -ne $DomainNode) { $DomainNode.InnerText } Else { "" }
            
            [Object]$UserIdNode = $DriveNode.SelectSingleNode("*[local-name()='UserId']")
            [String]$UserId = If ($Null -ne $UserIdNode) { $UserIdNode.InnerText } Else { "" }
            
            [Object]$AdGroupNode = $DriveNode.SelectSingleNode("*[local-name()='AdGroup']")
            [String]$AdGroup = If ($Null -ne $AdGroupNode) { $AdGroupNode.InnerText } Else { "" }
            
            [Object]$SiteNode = $DriveNode.SelectSingleNode("*[local-name()='Site']")
            [String]$Site = If ($Null -ne $SiteNode) { $SiteNode.InnerText } Else { "" }
            
            # Get DriveLetter using the correct element name "Drive" (as per VBScript)
            [Object]$DriveLetterNode = $DriveNode.SelectSingleNode("*[local-name()='Drive']")
            
            [String]$Drive = If ($Null -ne $DriveLetterNode) { $DriveLetterNode.InnerText } Else { "" }
            
            # Debug: Log what we're parsing
            [String]$DriveLetterText = If ($Null -ne $DriveLetterNode) { $DriveLetterNode.InnerText } Else { "NULL" }
            Write-DMLog "Mapper Drive: Parsing - DriveLetterNode: '$DriveLetterText', Drive: '$Drive'" -Level Verbose
            
            [Object]$UncPathNode = $DriveNode.SelectSingleNode("*[local-name()='UncPath']")
            [String]$UncPath = If ($Null -ne $UncPathNode) { $UncPathNode.InnerText } Else { "" }
            
            [Object]$DescriptionNode = $DriveNode.SelectSingleNode("*[local-name()='Description']")
            [String]$Description = If ($Null -ne $DescriptionNode) { $DescriptionNode.InnerText } Else { "" }
            
            [Object]$DisconnectNode = $DriveNode.SelectSingleNode("*[local-name()='DisconnectOnLogin']")
            [String]$DisconnectOnLogin = If ($Null -ne $DisconnectNode) { $DisconnectNode.InnerText } Else { "false" }
            
            $DriveObject = [PSCustomObject]@{
                PSTypeName = 'DM.DriveMapping'
                Id = $Id
                Domain = $Domain
                UserId = $UserId
                AdGroup = $AdGroup
                Site = $Site
                DriveLetter = $Drive
                UncPath = $UncPath
                Description = $Description
                DisconnectOnLogin = ($DisconnectOnLogin -eq "true")
            }
            
            # Log the constructed object details
            Write-DMLog "Mapper Drive: Constructed drive object:" -Level Verbose
            Write-DMLog "Mapper Drive:   DriveLetter: '$($DriveObject.DriveLetter)' (Empty: $([String]::IsNullOrEmpty($DriveObject.DriveLetter)))" -Level Verbose
            
            $Drives += $DriveObject
        }
        
        Return $Drives
    }
    Catch {
        Write-DMLog "Error parsing drive mappings: $($_.Exception.Message)" -Level Warning
        Return @()
    }
}

<#
.SYNOPSIS
    Parses printer mappings from SOAP response XML.
#>
Function Parse-DMPrinterMappingsResponse {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Xml]$ResponseXML
    )
    
    Try {
        [Array]$Printers = @()
        
        [Object]$ResultNode = Get-DMSOAPResult -ResponseXML $ResponseXML -MethodName "GetUserPrinters"
        
        If ($Null -eq $ResultNode) {
            Return @()
        }
        
        # Parse Printer elements - VBScript uses GetUserPrintersResult structure with MapperPrinter elements
        [Array]$PrinterNodes = $ResultNode.SelectNodes("./*[local-name()='MapperPrinter']")
        
        ForEach ($PrinterNode in $PrinterNodes) {
            [Object]$IdNode = $PrinterNode.SelectSingleNode("*[local-name()='Id']")
            [String]$Id = If ($Null -ne $IdNode) { $IdNode.InnerText } Else { "" }
            
            [Object]$UncPathNode = $PrinterNode.SelectSingleNode("*[local-name()='UncPath']")
            [String]$UncPath = If ($Null -ne $UncPathNode) { $UncPathNode.InnerText } Else { "" }
            
            [Object]$IsDefaultNode = $PrinterNode.SelectSingleNode("*[local-name()='IsDefault']")
            [String]$IsDefault = If ($Null -ne $IsDefaultNode) { $IsDefaultNode.InnerText } Else { "false" }
            
            [Object]$DescriptionNode = $PrinterNode.SelectSingleNode("*[local-name()='Description']")
            [String]$Description = If ($Null -ne $DescriptionNode) { $DescriptionNode.InnerText } Else { "" }
            
            $Printers += [PSCustomObject]@{
                PSTypeName = 'DM.PrinterMapping'
                Id = $Id
                UncPath = $UncPath
                IsDefault = ($IsDefault -eq "true")
                Description = $Description
            }
        }
        
        Return $Printers
    }
    Catch {
        Write-DMLog "Error parsing printer mappings: $($_.Exception.Message)" -Level Warning
        Return @()
    }
}

<#
.SYNOPSIS
    Parses PST mappings from SOAP response XML.
#>
Function Parse-DMPSTMappingsResponse {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Xml]$ResponseXML
    )
    
    Try {
        [Array]$PSTs = @()
        
        [Object]$ResultNode = Get-DMSOAPResult -ResponseXML $ResponseXML -MethodName "GetUserPersonalFolders"
        
        If ($Null -eq $ResultNode) {
            Return @()
        }
        
        # Parse PersonalFolder elements - VBScript uses GetUserPersonalFoldersResult structure with MapperPersonalFolder elements
        [Array]$PSTNodes = $ResultNode.SelectNodes("./*[local-name()='MapperPersonalFolder']")
        
        ForEach ($PSTNode in $PSTNodes) {
            [Object]$IdNode = $PSTNode.SelectSingleNode("*[local-name()='Id']")
            [String]$Id = If ($Null -ne $IdNode) { $IdNode.InnerText } Else { "" }
            
            [Object]$UserIdNode = $PSTNode.SelectSingleNode("*[local-name()='UserId']")
            [String]$UserId = If ($Null -ne $UserIdNode) { $UserIdNode.InnerText } Else { "" }
            
            [Object]$UncPathNode = $PSTNode.SelectSingleNode("*[local-name()='UncPath']")
            [String]$UncPath = If ($Null -ne $UncPathNode) { $UncPathNode.InnerText } Else { "" }
            
            [Object]$DisconnectNode = $PSTNode.SelectSingleNode("*[local-name()='DisconnectOnLogin']")
            [String]$DisconnectOnLogin = If ($Null -ne $DisconnectNode) { $DisconnectNode.InnerText } Else { "false" }
            
            $PSTs += [PSCustomObject]@{
                PSTypeName = 'DM.PSTMapping'
                Id = $Id
                UserId = $UserId
                UncPath = $UncPath
                DisconnectOnLogin = ($DisconnectOnLogin -eq "true")
            }
        }
        
        Return $PSTs
    }
    Catch {
        Write-DMLog "Error parsing PST mappings: $($_.Exception.Message)" -Level Warning
        Return @()
    }
}

<#
.SYNOPSIS
    Gets mapper service server.
    
.DESCRIPTION
    Discovers and returns the mapper service server.
    
.PARAMETER Domain
    Domain name
    
.OUTPUTS
    PSCustomObject - server information
    
.EXAMPLE
    $Server = Get-DMMapperServer -Domain $Computer.Domain
#>
Function Get-DMMapperServer {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$Domain = ""
    )
    
    Return Get-DMServiceServer -ServiceName "ClassicMapper.asmx" -Domain $Domain
}

# Export module members
Export-ModuleMember -Function @(
    'Get-DMDriveMappings',
    'Get-DMPrinterMappings',
    'Get-DMPSTMappings',
    'Get-DMMapperServer'
)

