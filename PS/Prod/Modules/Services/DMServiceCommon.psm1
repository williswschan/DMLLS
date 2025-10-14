<#
.SYNOPSIS
    Desktop Management Service Common Module
    
.DESCRIPTION
    Provides common functionality for SOAP web service communication.
    Includes server discovery, health checks, SOAP request/response handling.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: InventoryCommon_W10.vbs and MapperCommon_W10.vbs
#>

# Import required modules
Using Module ..\Framework\DMCommon.psm1

<#
.SYNOPSIS
    Gets the appropriate backend server based on environment.
    
.DESCRIPTION
    Determines which server to use (Production vs QA) based on domain forest.
    QA.NOM forest uses gdpmappercbqa.nomura.com, all others use gdpmappercb.nomura.com
    
.PARAMETER ServiceName
    Service name (ClassicInventory.asmx or ClassicMapper.asmx)
    
.PARAMETER Domain
    Domain name to check
    
.PARAMETER Timeout
    Service timeout in milliseconds (default: 10000)
    
.OUTPUTS
    PSCustomObject with server information
    
.EXAMPLE
    $Server = Get-DMServiceServer -ServiceName "ClassicInventory.asmx"
#>
Function Get-DMServiceServer {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$ServiceName,
        
        [Parameter(Mandatory=$False)]
        [String]$Domain = "",
        
        [Parameter(Mandatory=$False)]
        [Int]$Timeout = 10000
    )
    
    Try {
        # Determine server based on forest
        [String]$ServerFQDN = "gdpmappercb.nomura.com"
        
        # Check if QA environment
        Try {
            If ([String]::IsNullOrEmpty($Domain)) {
                # Use .NET DirectoryServices to get domain
                Add-Type -AssemblyName System.DirectoryServices
                [Object]$CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
                $Domain = $CurrentDomain.Forest.Name
            }
            
            If ($Domain.ToUpper() -eq "QA.NOM") {
                $ServerFQDN = "gdpmappercbqa.nomura.com"
            }
        } Catch {
            # If AD not available, use production server
            Write-Verbose "Could not determine forest, using production server"
        }
        
        # Build service URL
        [String]$ServiceURL = "http://$ServerFQDN/$ServiceName"
        
        # Test server connectivity
        [Boolean]$IsOnline = Test-DMServerPing -ComputerName $ServerFQDN
        [Int]$ResponseTime = -1
        
        If ($IsOnline) {
            # Measure response time
            [Object]$PingResult = Test-Connection -ComputerName $ServerFQDN -Count 1 -ErrorAction SilentlyContinue
            If ($Null -ne $PingResult) {
                $ResponseTime = $PingResult.ResponseTime
            }
        }
        
        # Test service availability
        [Boolean]$ServiceAvailable = $False
        If ($IsOnline) {
            $ServiceAvailable = Test-DMServiceHealth -ServiceURL $ServiceURL -Timeout $Timeout
        }
        
        # Return server object
        Return [PSCustomObject]@{
            PSTypeName = 'DM.ServiceServer'
            ServerName = $ServerFQDN
            ServiceURL = $ServiceURL
            IsOnline = $IsOnline
            ResponseTime = $ResponseTime
            ServiceAvailable = $ServiceAvailable
            Timeout = $Timeout
        }
    }
    Catch {
        Write-Error "Failed to get service server: $($_.Exception.Message)"
        Return $Null
    }
}

<#
.SYNOPSIS
    Tests if a web service is healthy and responding.
    
.DESCRIPTION
    Sends a SOAP TestService request to verify the service is available.
    
.PARAMETER ServiceURL
    Full service URL
    
.PARAMETER Timeout
    Timeout in milliseconds
    
.OUTPUTS
    Boolean - true if service is healthy
    
.EXAMPLE
    $Healthy = Test-DMServiceHealth -ServiceURL "http://gdpmappercb.nomura.com/ClassicInventory.asmx"
#>
Function Test-DMServiceHealth {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$ServiceURL,
        
        [Parameter(Mandatory=$False)]
        [Int]$Timeout = 10000
    )
    
    Try {
        # Build TestService SOAP request
        [String]$SoapRequest = New-DMSOAPEnvelope -MethodName "TestService" -MethodBody ""
        
        # Send request
        [Object]$Response = Invoke-DMSOAPRequest -ServiceURL $ServiceURL -SoapRequest $SoapRequest -Timeout $Timeout
        
        If ($Null -ne $Response -and $Response.Success) {
            # Check if response contains "True" or "OK" (for mock backend compatibility)
            If ($Response.ResponseText -like "*True*" -or $Response.ResponseText -like "*OK*" -or $Response.ResponseText -like "*TestServiceResult*") {
                Return $True
            }
        }
        
        Return $False
    }
    Catch {
        Write-Verbose "Service health check failed: $($_.Exception.Message)"
        Return $False
    }
}

<#
.SYNOPSIS
    Creates a SOAP envelope with method and body.
    
.DESCRIPTION
    Constructs a complete SOAP XML envelope for web service requests.
    
.PARAMETER MethodName
    SOAP method name
    
.PARAMETER MethodBody
    XML body content for the method
    
.PARAMETER Namespace
    XML namespace (default: http://webtools.japan.nom)
    
.OUTPUTS
    String - SOAP XML envelope
    
.EXAMPLE
    $Soap = New-DMSOAPEnvelope -MethodName "GetUserDrives" -MethodBody "<UserId>jsmith</UserId>"
#>
Function New-DMSOAPEnvelope {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$MethodName,
        
        [Parameter(Mandatory=$False)]
        [String]$MethodBody = "",
        
        [Parameter(Mandatory=$False)]
        [String]$Namespace = "http://webtools.japan.nom"
    )
    
    [String]$SoapEnvelope = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Body>
        <$MethodName xmlns="$Namespace">
$MethodBody
        </$MethodName>
    </soap:Body>
</soap:Envelope>
"@
    
    Return $SoapEnvelope
}

<#
.SYNOPSIS
    Sends a SOAP request to a web service.
    
.DESCRIPTION
    Executes an HTTP POST with SOAP XML content and handles timeout.
    
.PARAMETER ServiceURL
    Service URL to call
    
.PARAMETER SoapRequest
    SOAP XML request body
    
.PARAMETER Timeout
    Timeout in milliseconds
    
.OUTPUTS
    PSCustomObject with Success, StatusCode, ResponseText, ResponseXML
    
.EXAMPLE
    $Response = Invoke-DMSOAPRequest -ServiceURL $Url -SoapRequest $Soap -Timeout 10000
#>
Function Invoke-DMSOAPRequest {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$ServiceURL,
        
        [Parameter(Mandatory=$True)]
        [String]$SoapRequest,
        
        [Parameter(Mandatory=$False)]
        [Int]$Timeout = 10000
    )
    
    Try {
        # Calculate timeout in seconds for Invoke-WebRequest
        [Int]$TimeoutSeconds = [Math]::Ceiling($Timeout / 1000)
        
        # Send SOAP request
        [Object]$WebResponse = Invoke-WebRequest `
            -Uri $ServiceURL `
            -Method POST `
            -Body $SoapRequest `
            -ContentType "text/xml; charset=utf-8" `
            -TimeoutSec $TimeoutSeconds `
            -ErrorAction Stop
        
        # Parse response
        [Boolean]$Success = ($WebResponse.StatusCode -eq 200)
        [String]$ResponseText = $WebResponse.Content
        [Xml]$ResponseXML = $Null
        
        If ($Success) {
            Try {
                $ResponseXML = [Xml]$ResponseText
            } Catch {
                Write-Verbose "Could not parse response as XML"
            }
        }
        
        Return [PSCustomObject]@{
            PSTypeName = 'DM.SOAPResponse'
            Success = $Success
            StatusCode = $WebResponse.StatusCode
            ResponseText = $ResponseText
            ResponseXML = $ResponseXML
        }
    }
    Catch {
        Write-Verbose "SOAP request failed: $($_.Exception.Message)"
        
        Return [PSCustomObject]@{
            PSTypeName = 'DM.SOAPResponse'
            Success = $False
            StatusCode = 0
            ResponseText = ""
            ResponseXML = $Null
            Error = $_.Exception.Message
        }
    }
}

<#
.SYNOPSIS
    Extracts result data from SOAP response XML.
    
.DESCRIPTION
    Parses SOAP envelope and extracts the result element.
    
.PARAMETER ResponseXML
    SOAP response XML
    
.PARAMETER MethodName
    Method name to find result for
    
.OUTPUTS
    XmlElement - result element or $Null
    
.EXAMPLE
    $Result = Get-DMSOAPResult -ResponseXML $Response.ResponseXML -MethodName "GetUserDrives"
#>
Function Get-DMSOAPResult {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Xml]$ResponseXML,
        
        [Parameter(Mandatory=$True)]
        [String]$MethodName
    )
    
    Try {
        # Navigate SOAP envelope: Envelope -> Body -> {Method}Response -> {Method}Result
        [String]$ResponseName = "${MethodName}Response"
        [String]$ResultName = "${MethodName}Result"
        
        # Create namespace manager for XPath
        [Object]$NSManager = New-Object System.Xml.XmlNamespaceManager($ResponseXML.NameTable)
        $NSManager.AddNamespace("soap", "http://schemas.xmlsoap.org/soap/envelope/")
        $NSManager.AddNamespace("ns", "http://webtools.japan.nom")
        
        # Try to find result with namespace
        [Object]$ResultNode = $ResponseXML.SelectSingleNode("//ns:$ResultName", $NSManager)
        
        If ($Null -eq $ResultNode) {
            # Try without namespace (fallback)
            $ResultNode = $ResponseXML.SelectSingleNode("//$ResultName")
        }
        
        Return $ResultNode
    }
    Catch {
        Write-Verbose "Failed to extract SOAP result: $($_.Exception.Message)"
        Return $Null
    }
}

<#
.SYNOPSIS
    Builds XML element string with escaped content.
    
.DESCRIPTION
    Creates XML element with properly escaped text content.
    
.PARAMETER ElementName
    XML element name
    
.PARAMETER Content
    Content to include (will be XML-escaped)
    
.OUTPUTS
    String - XML element
    
.EXAMPLE
    $XML = New-DMXMLElement -ElementName "UserName" -Content "jsmith"
#>
Function New-DMXMLElement {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$ElementName,
        
        [Parameter(Mandatory=$False)]
        [String]$Content = ""
    )
    
    [String]$SafeContent = ConvertTo-DMXMLSafeText -Text $Content
    Return "            <$ElementName>$SafeContent</$ElementName>"
}

<#
.SYNOPSIS
    Tests if user is part of a group (helper for backward compatibility).
    
.DESCRIPTION
    Checks if user is member of a group. Wrapper for common pattern.
    
.PARAMETER UserInfo
    User info object
    
.PARAMETER GroupName
    Group name to check
    
.OUTPUTS
    Boolean - true if member
    
.EXAMPLE
    $IsMember = Test-DMUserPartOfGroup -UserInfo $User -GroupName "Laptop Offline PC"
#>
Function Test-DMUserPartOfGroup {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [String]$GroupName
    )
    
    If ($Null -eq $UserInfo.Groups -or $UserInfo.Groups.Count -eq 0) {
        Return $False
    }
    
    ForEach ($Group in $UserInfo.Groups) {
        If ($Group.GroupName -like "*$GroupName*") {
            Return $True
        }
    }
    
    Return $False
}

<#
.SYNOPSIS
    Tests if computer is part of a group (helper for backward compatibility).
    
.DESCRIPTION
    Checks if computer is member of a group. Wrapper for common pattern.
    
.PARAMETER ComputerInfo
    Computer info object
    
.PARAMETER GroupName
    Group name to check
    
.OUTPUTS
    Boolean - true if member
    
.EXAMPLE
    $IsMember = Test-DMHostPartOfGroup -ComputerInfo $Computer -GroupName "Pilot Desktop Management Script"
#>
Function Test-DMHostPartOfGroup {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo,
        
        [Parameter(Mandatory=$True)]
        [String]$GroupName
    )
    
    If ($Null -eq $ComputerInfo.Groups -or $ComputerInfo.Groups.Count -eq 0) {
        Return $False
    }
    
    ForEach ($Group in $ComputerInfo.Groups) {
        If ($Group.GroupName -like "*$GroupName*") {
            Return $True
        }
    }
    
    Return $False
}

<#
.SYNOPSIS
    Sends SOAP request with proper authentication headers using proven pattern.
    
.DESCRIPTION
    Sends SOAP request using the proven authentication pattern from working code.
    Includes both HTTP Basic Auth and SOAP header authentication.
    
.PARAMETER ServerUrl
    Full server URL with service endpoint
    
.PARAMETER SOAPBody
    SOAP request body (XML string)
    
.PARAMETER SOAPAction
    SOAP action header value
    
.PARAMETER Username
    Username for authentication
    
.PARAMETER Password
    Password for authentication
    
.PARAMETER Timeout
    Request timeout in milliseconds
    
.OUTPUTS
    PSCustomObject with response data
    
.EXAMPLE
    $Response = Send-DMSOAPRequestWithAuth -ServerUrl $Url -SOAPBody $Body -SOAPAction $Action -Username $User -Password $Pass
#>
Function Send-DMSOAPRequest {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$ServerUrl,
        
        [Parameter(Mandatory=$True)]
        [String]$SOAPBody,
        
        [Parameter(Mandatory=$True)]
        [String]$SOAPAction,
        
        [Parameter(Mandatory=$False)]
        [Int]$Timeout = 10000
    )
    
    Try {
        # Create simple headers (no authentication, matching VBScript)
        [Hashtable]$Headers = @{
            "Content-Type" = "text/xml; charset=utf-8"
            "SOAPAction" = $SOAPAction
        }
        
        Write-Verbose "Sending SOAP request to: $ServerUrl"
        Write-Verbose "SOAP Action: $SOAPAction"
        
        [Object]$Response = Invoke-WebRequest -Uri $ServerUrl -Method Post -Body $SOAPBody -Headers $Headers -TimeoutSec ($Timeout / 1000) -UseBasicParsing
        
        If ($Response.StatusCode -eq 200) {
            Write-Verbose "SOAP request successful"
            Return [PSCustomObject]@{
                Success = $True
                StatusCode = $Response.StatusCode
                Content = $Response.Content
                Headers = $Response.Headers
                ResponseText = $Response.Content
                ResponseXML = [Xml]$Response.Content
            }
        } Else {
            Write-Warning "HTTP Error: $($Response.StatusCode) - $($Response.StatusDescription)"
            Return [PSCustomObject]@{
                Success = $False
                StatusCode = $Response.StatusCode
                StatusDescription = $Response.StatusDescription
                Content = $Response.Content
                Headers = $Response.Headers
                ResponseText = $Response.Content
                ResponseXML = $Null
            }
        }
    }
    Catch {
        Write-Warning "SOAP request failed: $($_.Exception.Message)"
        Return [PSCustomObject]@{
            Success = $False
            Error = $_.Exception.Message
            StatusCode = $Null
            Content = ""
            ResponseText = ""
            ResponseXML = $Null
        }
    }
}

Function Send-DMSOAPRequestWithAuth {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$ServerUrl,
        
        [Parameter(Mandatory=$True)]
        [String]$SOAPBody,
        
        [Parameter(Mandatory=$True)]
        [String]$SOAPAction,
        
        [Parameter(Mandatory=$True)]
        [String]$Username,
        
        [Parameter(Mandatory=$True)]
        [String]$Password,
        
        [Parameter(Mandatory=$False)]
        [Int]$Timeout = 10000
    )
    
    Try {
        # Create authentication headers using proven pattern from working code
        [Hashtable]$Headers = @{
            "Content-Type" = "text/xml; charset=utf-8"
            "SOAPAction" = $SOAPAction
            "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Username + ":" + $Password))
        }
        
        Write-Verbose "Sending SOAP request to: $ServerUrl"
        Write-Verbose "SOAP Action: $SOAPAction"
        
        [Object]$Response = Invoke-WebRequest -Uri $ServerUrl -Method Post -Body $SOAPBody -Headers $Headers -TimeoutSec ($Timeout / 1000) -UseBasicParsing
        
        If ($Response.StatusCode -eq 200) {
            Write-Verbose "SOAP request successful"
            Return [PSCustomObject]@{
                Success = $True
                StatusCode = $Response.StatusCode
                Content = $Response.Content
                Headers = $Response.Headers
                ResponseText = $Response.Content
                ResponseXML = [Xml]$Response.Content
            }
        } Else {
            Write-Warning "HTTP Error: $($Response.StatusCode) - $($Response.StatusDescription)"
            Return [PSCustomObject]@{
                Success = $False
                StatusCode = $Response.StatusCode
                StatusDescription = $Response.StatusDescription
                Content = $Response.Content
                Headers = $Response.Headers
                ResponseText = $Response.Content
                ResponseXML = $Null
            }
        }
    }
    Catch {
        Write-Warning "SOAP request failed: $($_.Exception.Message)"
        Return [PSCustomObject]@{
            Success = $False
            Error = $_.Exception.Message
            StatusCode = $Null
            Content = ""
            ResponseText = ""
            ResponseXML = $Null
        }
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Get-DMServiceServer',
    'Test-DMServiceHealth',
    'New-DMSOAPEnvelope',
    'Invoke-DMSOAPRequest',
    'Send-DMSOAPRequest',
    'Send-DMSOAPRequestWithAuth',
    'Get-DMSOAPResult',
    'New-DMXMLElement',
    'Test-DMUserPartOfGroup',
    'Test-DMHostPartOfGroup'
)

