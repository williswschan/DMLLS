<#
.SYNOPSIS
    Test the Actual Available Methods from WSDL
    
.DESCRIPTION
    Tests the methods that actually exist in the WSDL:
    - TestService
    - GetUserId  
    - GetDomains
    
.EXAMPLE
    .\Test-ActualMethods.ps1
#>

[CmdletBinding()]
Param()

Function Write-TestResult {
    Param(
        [String]$Section,
        [String]$Message,
        [String]$Value = "",
        [String]$Color = "White"
    )
    
    If ($Value) {
        Write-Host "[$Section] $Message : " -NoNewline -ForegroundColor Cyan
        Write-Host $Value -ForegroundColor $Color
    } Else {
        Write-Host "[$Section] $Message" -ForegroundColor $Color
    }
}

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "TESTING ACTUAL AVAILABLE METHODS" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Green

# Get current user info
[String]$Username = $env:USERNAME
[String]$ComputerName = $env:COMPUTERNAME

Write-TestResult "INFO" "Testing User" $Username
Write-TestResult "INFO" "Testing Computer" $ComputerName
Write-TestResult "INFO" "Server" "gdpmappercb.nomura.com"
Write-TestResult "INFO" "Namespace" "http://webtools.japan.nom"

# Test 1: TestService (No parameters, returns boolean)
Write-Host "`n" -NoNewline
Write-TestResult "TEST 1" "Testing TestService Method" -Color Yellow

[String]$TestService_SOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:web="http://webtools.japan.nom">
    <soap:Body>
        <web:TestService />
    </soap:Body>
</soap:Envelope>
"@

[Hashtable]$TestService_Headers = @{
    "Content-Type" = "text/xml; charset=utf-8"
    "SOAPAction" = "http://webtools.japan.nom/TestService"
}

Write-Host "SOAP Request:" -ForegroundColor White
Write-Host $TestService_SOAP -ForegroundColor Gray

Try {
    [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $TestService_SOAP -Headers $TestService_Headers -TimeoutSec 10 -UseBasicParsing
    
    Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
    Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
    
    Write-Host "`n--- TestService Response ---" -ForegroundColor Green
    Write-Host $Response.Content -ForegroundColor Gray
    
    # Parse response
    Try {
        [Xml]$ResponseXML = $Response.Content
        [Object]$TestServiceResult = $ResponseXML.SelectSingleNode("//*[local-name()='TestServiceResult']")
        
        If ($TestServiceResult) {
            Write-TestResult "SUCCESS" "TestService Result" $TestServiceResult.InnerText -Color Green
        } Else {
            Write-TestResult "WARNING" "Could not parse TestServiceResult" -Color Yellow
        }
    } Catch {
        Write-TestResult "ERROR" "Failed to parse XML response" $_.Exception.Message -Color Red
    }
    
} Catch {
    [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
    [String]$StatusDesc = if ($_.Exception.Response) { $_.Exception.Response.StatusDescription } else { "Unknown" }
    
    Write-TestResult "FAILED" "HTTP Status" "$StatusCode - $StatusDesc" -Color Red
    
    # Show error response
    Try {
        [Object]$ErrorStream = $_.Exception.Response.GetResponseStream()
        [Object]$Reader = New-Object System.IO.StreamReader($ErrorStream)
        [String]$ErrorBody = $Reader.ReadToEnd()
        
        If ($ErrorBody.Length -gt 0) {
            Write-Host "`n--- Error Response ---" -ForegroundColor Red
            Write-Host $ErrorBody -ForegroundColor Gray
        }
    } Catch {
        Write-TestResult "INFO" "Could not read error response" -Color Yellow
    }
}

# Test 2: GetUserId (No parameters, returns string + optional fault)
Write-Host "`n" -NoNewline
Write-TestResult "TEST 2" "Testing GetUserId Method" -Color Yellow

[String]$GetUserId_SOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:web="http://webtools.japan.nom">
    <soap:Body>
        <web:GetUserId />
    </soap:Body>
</soap:Envelope>
"@

[Hashtable]$GetUserId_Headers = @{
    "Content-Type" = "text/xml; charset=utf-8"
    "SOAPAction" = "http://webtools.japan.nom/GetUserId"
}

Write-Host "SOAP Request:" -ForegroundColor White
Write-Host $GetUserId_SOAP -ForegroundColor Gray

Try {
    [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $GetUserId_SOAP -Headers $GetUserId_Headers -TimeoutSec 10 -UseBasicParsing
    
    Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
    Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
    
    Write-Host "`n--- GetUserId Response ---" -ForegroundColor Green
    Write-Host $Response.Content -ForegroundColor Gray
    
    # Parse response
    Try {
        [Xml]$ResponseXML = $Response.Content
        [Object]$GetUserIdResult = $ResponseXML.SelectSingleNode("//*[local-name()='GetUserIdResult']")
        [Object]$Fault = $ResponseXML.SelectSingleNode("//*[local-name()='Fault']")
        
        If ($GetUserIdResult) {
            Write-TestResult "SUCCESS" "GetUserId Result" $GetUserIdResult.InnerText -Color Green
        } Else {
            Write-TestResult "WARNING" "Could not parse GetUserIdResult" -Color Yellow
        }
        
        If ($Fault) {
            Write-TestResult "WARNING" "Fault detected" -Color Yellow
            [Object]$FaultType = $Fault.SelectSingleNode("*[local-name()='FaultType']")
            [Object]$Message = $Fault.SelectSingleNode("*[local-name()='Message']")
            [Object]$Source = $Fault.SelectSingleNode("*[local-name()='Source']")
            
            If ($FaultType) { Write-TestResult "FAULT" "FaultType" $FaultType.InnerText -Color Red }
            If ($Message) { Write-TestResult "FAULT" "Message" $Message.InnerText -Color Red }
            If ($Source) { Write-TestResult "FAULT" "Source" $Source.InnerText -Color Red }
        }
    } Catch {
        Write-TestResult "ERROR" "Failed to parse XML response" $_.Exception.Message -Color Red
    }
    
} Catch {
    [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
    [String]$StatusDesc = if ($_.Exception.Response) { $_.Exception.Response.StatusDescription } else { "Unknown" }
    
    Write-TestResult "FAILED" "HTTP Status" "$StatusCode - $StatusDesc" -Color Red
    
    # Show error response
    Try {
        [Object]$ErrorStream = $_.Exception.Response.GetResponseStream()
        [Object]$Reader = New-Object System.IO.StreamReader($ErrorStream)
        [String]$ErrorBody = $Reader.ReadToEnd()
        
        If ($ErrorBody.Length -gt 0) {
            Write-Host "`n--- Error Response ---" -ForegroundColor Red
            Write-Host $ErrorBody -ForegroundColor Gray
        }
    } Catch {
        Write-TestResult "INFO" "Could not read error response" -Color Yellow
    }
}

# Test 3: GetDomains (No parameters, returns ArrayOfMapperDomain + optional fault)
Write-Host "`n" -NoNewline
Write-TestResult "TEST 3" "Testing GetDomains Method" -Color Yellow

[String]$GetDomains_SOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:web="http://webtools.japan.nom">
    <soap:Body>
        <web:GetDomains />
    </soap:Body>
</soap:Envelope>
"@

[Hashtable]$GetDomains_Headers = @{
    "Content-Type" = "text/xml; charset=utf-8"
    "SOAPAction" = "http://webtools.japan.nom/GetDomains"
}

Write-Host "SOAP Request:" -ForegroundColor White
Write-Host $GetDomains_SOAP -ForegroundColor Gray

Try {
    [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $GetDomains_SOAP -Headers $GetDomains_Headers -TimeoutSec 10 -UseBasicParsing
    
    Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
    Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
    
    Write-Host "`n--- GetDomains Response ---" -ForegroundColor Green
    Write-Host $Response.Content -ForegroundColor Gray
    
    # Parse response
    Try {
        [Xml]$ResponseXML = $Response.Content
        [Object]$GetDomainsResult = $ResponseXML.SelectSingleNode("//*[local-name()='GetDomainsResult']")
        [Object]$Fault = $ResponseXML.SelectSingleNode("//*[local-name()='Fault']")
        
        If ($GetDomainsResult) {
            Write-TestResult "SUCCESS" "GetDomains Result found" -Color Green
            
            # Look for MapperDomain objects
            [Array]$MapperDomains = $ResponseXML.SelectNodes("//*[local-name()='MapperDomain']")
            Write-TestResult "INFO" "MapperDomain objects found" $MapperDomains.Count -Color Cyan
            
            If ($MapperDomains.Count -gt 0) {
                Write-Host "`n--- Domain Details ---" -ForegroundColor Cyan
                ForEach ($Domain in $MapperDomains) {
                    [Object]$Id = $Domain.SelectSingleNode("*[local-name()='Id']")
                    [Object]$ParentId = $Domain.SelectSingleNode("*[local-name()='ParentId']")
                    [Object]$OuMapping = $Domain.SelectSingleNode("*[local-name()='OuMapping']")
                    [Object]$Name = $Domain.SelectSingleNode("*[local-name()='Name']")
                    [Object]$DomainName = $Domain.SelectSingleNode("*[local-name()='Domain']")
                    
                    Write-Host "  Domain:" -ForegroundColor White
                    If ($Id) { Write-Host "    Id: $($Id.InnerText)" -ForegroundColor Gray }
                    If ($ParentId) { Write-Host "    ParentId: $($ParentId.InnerText)" -ForegroundColor Gray }
                    If ($OuMapping) { Write-Host "    OuMapping: $($OuMapping.InnerText)" -ForegroundColor Gray }
                    If ($Name) { Write-Host "    Name: $($Name.InnerText)" -ForegroundColor Gray }
                    If ($DomainName) { Write-Host "    Domain: $($DomainName.InnerText)" -ForegroundColor Gray }
                    Write-Host ""
                }
            }
        } Else {
            Write-TestResult "WARNING" "Could not parse GetDomainsResult" -Color Yellow
        }
        
        If ($Fault) {
            Write-TestResult "WARNING" "Fault detected" -Color Yellow
            [Object]$FaultType = $Fault.SelectSingleNode("*[local-name()='FaultType']")
            [Object]$Message = $Fault.SelectSingleNode("*[local-name()='Message']")
            [Object]$Source = $Fault.SelectSingleNode("*[local-name()='Source']")
            
            If ($FaultType) { Write-TestResult "FAULT" "FaultType" $FaultType.InnerText -Color Red }
            If ($Message) { Write-TestResult "FAULT" "Message" $Message.InnerText -Color Red }
            If ($Source) { Write-TestResult "FAULT" "Source" $Source.InnerText -Color Red }
        }
    } Catch {
        Write-TestResult "ERROR" "Failed to parse XML response" $_.Exception.Message -Color Red
    }
    
} Catch {
    [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
    [String]$StatusDesc = if ($_.Exception.Response) { $_.Exception.Response.StatusDescription } else { "Unknown" }
    
    Write-TestResult "FAILED" "HTTP Status" "$StatusCode - $StatusDesc" -Color Red
    
    # Show error response
    Try {
        [Object]$ErrorStream = $_.Exception.Response.GetResponseStream()
        [Object]$Reader = New-Object System.IO.StreamReader($ErrorStream)
        [String]$ErrorBody = $Reader.ReadToEnd()
        
        If ($ErrorBody.Length -gt 0) {
            Write-Host "`n--- Error Response ---" -ForegroundColor Red
            Write-Host $ErrorBody -ForegroundColor Gray
        }
    } Catch {
        Write-TestResult "INFO" "Could not read error response" -Color Yellow
    }
}

# Test 4: Try with Authentication (in case it's required)
Write-Host "`n" -NoNewline
Write-TestResult "TEST 4" "Testing with Authentication" -Color Yellow

[String]$Auth_SOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:web="http://webtools.japan.nom">
    <soap:Header>
        <web:AuthHeader>
            <web:Username>$Username</web:Username>
            <web:Password>placeholder</web:Password>
        </web:AuthHeader>
    </soap:Header>
    <soap:Body>
        <web:TestService />
    </soap:Body>
</soap:Envelope>
"@

[Hashtable]$Auth_Headers = @{
    "Content-Type" = "text/xml; charset=utf-8"
    "SOAPAction" = "http://webtools.japan.nom/TestService"
    "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$Username`:placeholder"))
}

Write-Host "SOAP Request with Auth:" -ForegroundColor White
Write-Host $Auth_SOAP -ForegroundColor Gray

Try {
    [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $Auth_SOAP -Headers $Auth_Headers -TimeoutSec 10 -UseBasicParsing
    
    Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
    Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
    
    Write-Host "`n--- Authenticated TestService Response ---" -ForegroundColor Green
    Write-Host $Response.Content -ForegroundColor Gray
    
} Catch {
    [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
    [String]$StatusDesc = if ($_.Exception.Response) { $_.Exception.Response.StatusDescription } else { "Unknown" }
    
    Write-TestResult "FAILED" "HTTP Status" "$StatusCode - $StatusDesc" -Color Red
}

# Summary
Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "ANALYSIS SUMMARY" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

Write-Host "`nKEY FINDINGS:" -ForegroundColor Yellow
Write-Host "1. This service only has 3 methods: TestService, GetUserId, GetDomains" -ForegroundColor White
Write-Host "2. The drive/printer/PST mapping methods do NOT exist on this service" -ForegroundColor White
Write-Host "3. This appears to be a different service than expected" -ForegroundColor White

Write-Host "`nNEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. Check if there are other services (ClassicMapper2.asmx, DriveMapper.asmx, etc.)" -ForegroundColor White
Write-Host "2. Look for the actual mapping service endpoints" -ForegroundColor White
Write-Host "3. Check if the API has changed or methods were moved" -ForegroundColor White
Write-Host "4. Verify this is the correct server for mapping operations" -ForegroundColor White

Write-Host "`nPOSSIBLE SOLUTIONS:" -ForegroundColor Yellow
Write-Host "1. Find the correct service endpoint for mappings" -ForegroundColor White
Write-Host "2. Check if mappings are handled differently (REST API, different SOAP service)" -ForegroundColor White
Write-Host "3. Verify the service URL with your backend team" -ForegroundColor White

Write-Host "`n"
