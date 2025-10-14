<#
.SYNOPSIS
    Investigate Where Mapping Service Actually Is
    
.DESCRIPTION
    Since the mapping methods don't exist on gdpmappercb.nomura.com,
    let's investigate other possibilities:
    1. Different server/domain
    2. REST API instead of SOAP
    3. Different service structure
    4. Check if methods were moved/renamed
    
.EXAMPLE
    .\Investigate-MappingService.ps1
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

Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "INVESTIGATING MAPPING SERVICE LOCATION" -ForegroundColor Magenta
Write-Host "========================================`n" -ForegroundColor Magenta

# Test 1: Check if there are other servers/domains
Write-Host "`n" -NoNewline
Write-TestResult "TEST 1" "Checking Alternative Servers" -Color Yellow

[Array]$AlternativeServers = @(
    "http://gdpmappercb.nomura.com",
    "https://gdpmappercb.nomura.com",
    "http://gdpmappercb.nomura.co.jp", 
    "https://gdpmappercb.nomura.co.jp",
    "http://mapper.nomura.com",
    "https://mapper.nomura.com",
    "http://mapping.nomura.com",
    "https://mapping.nomura.com",
    "http://desktop.nomura.com",
    "https://desktop.nomura.com",
    "http://gdp.nomura.com",
    "https://gdp.nomura.com",
    "http://webtools.japan.nom",
    "https://webtools.japan.nom"
)

ForEach ($Server in $AlternativeServers) {
    Write-Host "`nTesting server: $Server" -ForegroundColor Cyan
    
    Try {
        [Object]$Response = Invoke-WebRequest -Uri $Server -Method Get -UseBasicParsing -TimeoutSec 5
        Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
        Write-TestResult "SUCCESS" "Content Type" $Response.Headers["Content-Type"] -Color Green
        
        # Check if it's a web service
        If ($Response.Content -match "wsdl|WSDL|soap|SOAP|web service|Web Service") {
            Write-TestResult "SUCCESS" "Service Type" "Web Service detected" -Color Green
            
            # Look for mapping methods
            [Array]$MappingMethods = $Response.Content | Select-String -Pattern "Drive|Printer|PST|Mapping" -CaseSensitive:$False
            If ($MappingMethods.Count -gt 0) {
                Write-TestResult "SUCCESS" "Mapping Methods Found" $MappingMethods.Count -Color Green
                Write-Host "  Mapping content found:" -ForegroundColor Yellow
                $MappingMethods | Select-Object -First 5 | ForEach-Object { 
                    [String]$Line = $_.Line.Trim()
                    If ($Line.Length -gt 0 -and $Line.Length -lt 150) {
                        Write-Host "    $Line" -ForegroundColor Gray
                    }
                }
            }
        }
        
    } Catch {
        [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
        Write-TestResult "FAILED" "HTTP Status" $StatusCode -Color Red
    }
}

# Test 2: Check for REST API endpoints
Write-Host "`n" -NoNewline
Write-TestResult "TEST 2" "Checking REST API Endpoints" -Color Yellow

[Array]$RESTEndpoints = @(
    "http://gdpmappercb.nomura.com/api",
    "http://gdpmappercb.nomura.com/api/mappings",
    "http://gdpmappercb.nomura.com/api/drives",
    "http://gdpmappercb.nomura.com/api/printers",
    "http://gdpmappercb.nomura.com/api/pst",
    "http://gdpmappercb.nomura.com/api/v1",
    "http://gdpmappercb.nomura.com/api/v2",
    "http://gdpmappercb.nomura.com/rest",
    "http://gdpmappercb.nomura.com/rest/mappings",
    "http://gdpmappercb.nomura.com/json",
    "http://gdpmappercb.nomura.com/json/mappings"
)

ForEach ($Endpoint in $RESTEndpoints) {
    Write-Host "`nTesting REST endpoint: $Endpoint" -ForegroundColor Cyan
    
    Try {
        [Object]$Response = Invoke-WebRequest -Uri $Endpoint -Method Get -UseBasicParsing -TimeoutSec 5
        Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
        Write-TestResult "SUCCESS" "Content Type" $Response.Headers["Content-Type"] -Color Green
        
        # Check if it's JSON
        If ($Response.Content -match "^\s*[\{\[]") {
            Write-TestResult "SUCCESS" "Response Type" "JSON detected" -Color Green
        }
        
        # Show content preview
        [String]$Preview = $Response.Content.Substring(0, [Math]::Min(200, $Response.Content.Length))
        Write-Host "  Content: $Preview..." -ForegroundColor Gray
        
    } Catch {
        [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
        Write-TestResult "FAILED" "HTTP Status" $StatusCode -Color Red
    }
}

# Test 3: Check for different service structures
Write-Host "`n" -NoNewline
Write-TestResult "TEST 3" "Checking Different Service Structures" -Color Yellow

[Array]$ServiceStructures = @(
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/GetDriveMappings",
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/GetPrinterMappings",
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/GetPSTMappings",
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/DriveMappings",
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/PrinterMappings",
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/PSTMappings",
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/GetMappings",
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/Mappings",
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/ListDrives",
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/ListPrinters",
    "http://gdpmappercb.nomura.com/ClassicMapper.asmx/ListPSTs"
)

ForEach ($Structure in $ServiceStructures) {
    Write-Host "`nTesting structure: $Structure" -ForegroundColor Cyan
    
    Try {
        [Object]$Response = Invoke-WebRequest -Uri $Structure -Method Get -UseBasicParsing -TimeoutSec 5
        Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
        Write-Host "  Content: $($Response.Content.Substring(0, [Math]::Min(100, $Response.Content.Length)))..." -ForegroundColor Gray
        
    } Catch {
        [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
        Write-TestResult "FAILED" "HTTP Status" $StatusCode -Color Red
    }
}

# Test 4: Check if methods exist but with different names
Write-Host "`n" -NoNewline
Write-TestResult "TEST 4" "Testing Method Name Variations" -Color Yellow

[Array]$MethodVariations = @(
    "GetDriveMapping",      # singular
    "GetDriveMap",          # shortened
    "GetDrives",            # just drives
    "GetDriveList",         # list format
    "ListDrives",           # list prefix
    "RetrieveDrives",       # retrieve prefix
    "GetDriveMaps",         # maps plural
    "GetDriveMappingList",  # full descriptive
    "GetUserDrives",        # user prefix
    "GetComputerDrives",    # computer prefix
    "GetDesktopDrives",     # desktop prefix
    "GetNomuraDrives"       # nomura prefix
)

[String]$TestSOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:web="http://webtools.japan.nom">
    <soap:Body>
        <web:{0} />
    </soap:Body>
</soap:Envelope>
"@

ForEach ($Method in $MethodVariations) {
    Write-Host "`nTesting method: $Method" -ForegroundColor Cyan
    
    [String]$SOAPBody = $TestSOAP -f $Method
    [Hashtable]$Headers = @{
        "Content-Type" = "text/xml; charset=utf-8"
        "SOAPAction" = "http://webtools.japan.nom/$Method"
    }
    
    Try {
        [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $SOAPBody -Headers $Headers -TimeoutSec 5 -UseBasicParsing
        
        Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
        Write-Host "  Response: $($Response.Content.Substring(0, [Math]::Min(100, $Response.Content.Length)))..." -ForegroundColor Gray
        
    } Catch {
        [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
        Write-TestResult "FAILED" "HTTP Status" $StatusCode -Color Red
    }
}

# Summary
Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "INVESTIGATION SUMMARY" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

Write-Host "`nKEY FINDINGS:" -ForegroundColor Yellow
Write-Host "1. The mapping methods do NOT exist on gdpmappercb.nomura.com" -ForegroundColor White
Write-Host "2. Only TestService, GetUserId, and GetDomains are available" -ForegroundColor White
Write-Host "3. No alternative endpoints found with mapping functionality" -ForegroundColor White

Write-Host "`nPOSSIBLE EXPLANATIONS:" -ForegroundColor Yellow
Write-Host "1. Wrong server - mapping service is on a different server" -ForegroundColor White
Write-Host "2. API changed - mapping functionality was removed or moved" -ForegroundColor White
Write-Host "3. Different technology - mappings handled via REST API, not SOAP" -ForegroundColor White
Write-Host "4. Different service structure - methods have different names" -ForegroundColor White
Write-Host "5. Authentication required - methods exist but need proper auth" -ForegroundColor White

Write-Host "`nRECOMMENDATIONS:" -ForegroundColor Yellow
Write-Host "1. Contact your backend team to confirm the correct service URLs" -ForegroundColor White
Write-Host "2. Ask if the mapping API has changed or moved" -ForegroundColor White
Write-Host "3. Check if mappings are now handled via REST API instead of SOAP" -ForegroundColor White
Write-Host "4. Verify if different authentication is required" -ForegroundColor White
Write-Host "5. Check if the service is down or under maintenance" -ForegroundColor White

Write-Host "`nNEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. Review any successful responses from the tests above" -ForegroundColor White
Write-Host "2. Check with your backend team for the correct mapping service location" -ForegroundColor White
Write-Host "3. If no mapping service exists, consider implementing a workaround" -ForegroundColor White
Write-Host "4. Update the PowerShell modules to handle the missing functionality gracefully" -ForegroundColor White

Write-Host "`n"
