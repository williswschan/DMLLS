<#
.SYNOPSIS
    Test for Other Possible Service Endpoints
    
.DESCRIPTION
    Tests various possible service endpoints that might contain
    the drive/printer/PST mapping methods.
    
.EXAMPLE
    .\Test-OtherEndpoints.ps1
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
Write-Host "TESTING OTHER POSSIBLE ENDPOINTS" -ForegroundColor Magenta
Write-Host "========================================`n" -ForegroundColor Magenta

# List of possible endpoints to test
[Array]$Endpoints = @(
    @{ Name = "ClassicMapper2.asmx"; Description = "Alternative mapper service" },
    @{ Name = "DriveMapper.asmx"; Description = "Drive-specific mapper" },
    @{ Name = "PrinterMapper.asmx"; Description = "Printer-specific mapper" },
    @{ Name = "PSTMapper.asmx"; Description = "PST-specific mapper" },
    @{ Name = "Mapper.asmx"; Description = "Generic mapper service" },
    @{ Name = "MappingService.asmx"; Description = "Full mapping service" },
    @{ Name = "DesktopMapper.asmx"; Description = "Desktop mapping service" },
    @{ Name = "UserMapper.asmx"; Description = "User mapping service" },
    @{ Name = "ResourceMapper.asmx"; Description = "Resource mapping service" },
    @{ Name = "GDPMapper.asmx"; Description = "GDP-specific mapper" },
    @{ Name = "NomuraMapper.asmx"; Description = "Nomura-specific mapper" },
    @{ Name = "WebMapper.asmx"; Description = "Web mapper service" },
    @{ Name = "Service.asmx"; Description = "Generic service" },
    @{ Name = "API.asmx"; Description = "API service" },
    @{ Name = "WebService.asmx"; Description = "Generic web service" }
)

[String]$BaseURL = "http://gdpmappercb.nomura.com"
[Int]$FoundCount = 0
[Int]$TestCount = 0

Write-Host "Testing $($Endpoints.Count) possible endpoints..." -ForegroundColor White

ForEach ($Endpoint in $Endpoints) {
    $TestCount++
    [String]$EndpointName = $Endpoint.Name
    [String]$Description = $Endpoint.Description
    [String]$FullURL = "$BaseURL/$EndpointName"
    
    Write-Host "`n[$TestCount] Testing: $EndpointName" -ForegroundColor Cyan
    Write-Host "  Description: $Description" -ForegroundColor Gray
    Write-Host "  URL: $FullURL" -ForegroundColor Gray
    
    Try {
        # Test basic connectivity
        [Object]$Response = Invoke-WebRequest -Uri $FullURL -Method Get -UseBasicParsing -TimeoutSec 5
        
        Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
        Write-TestResult "SUCCESS" "Content Type" $Response.Headers["Content-Type"] -Color Green
        Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
        
        # Check if it's a web service
        If ($Response.Content -match "wsdl|WSDL|soap|SOAP|web service|Web Service") {
            Write-TestResult "SUCCESS" "Service Type" "Web Service detected" -Color Green
            
            # Try to get WSDL
            Try {
                [Object]$WSDLResponse = Invoke-WebRequest -Uri "$FullURL?WSDL" -UseBasicParsing -TimeoutSec 5
                Write-TestResult "SUCCESS" "WSDL Available" "Yes" -Color Green
                
                # Look for mapping methods in WSDL
                [Array]$MappingMethods = $WSDLResponse.Content | Select-String -Pattern "Drive|Printer|PST|Mapping" -CaseSensitive:$False
                If ($MappingMethods.Count -gt 0) {
                    Write-TestResult "SUCCESS" "Mapping Methods Found" $MappingMethods.Count -Color Green
                    Write-Host "  Mapping-related content:" -ForegroundColor Yellow
                    $MappingMethods | ForEach-Object { 
                        [String]$Line = $_.Line.Trim()
                        If ($Line.Length -gt 0 -and $Line.Length -lt 200) {
                            Write-Host "    $Line" -ForegroundColor Gray
                        }
                    }
                } Else {
                    Write-TestResult "INFO" "No mapping methods found in WSDL" -Color Yellow
                }
                
                # Save WSDL for inspection
                [String]$WSDLPath = "C:\Temp\$EndpointName.wsdl"
                $WSDLResponse.Content | Out-File -FilePath $WSDLPath -Encoding UTF8
                Write-TestResult "INFO" "WSDL saved to" $WSDLPath -Color Cyan
                
            } Catch {
                Write-TestResult "WARNING" "WSDL not available" -Color Yellow
            }
            
            $FoundCount++
            
        } Else {
            Write-TestResult "INFO" "Service Type" "Not a web service (HTML page)" -Color Yellow
        }
        
        # Show content preview
        [String]$Preview = $Response.Content.Substring(0, [Math]::Min(200, $Response.Content.Length))
        Write-Host "  Content Preview: $Preview..." -ForegroundColor Gray
        
    } Catch {
        [Int]$StatusCode = 0
        [String]$StatusDesc = "Unknown"
        
        If ($Null -ne $_.Exception.Response) {
            $StatusCode = [Int]$_.Exception.Response.StatusCode.value__
            $StatusDesc = $_.Exception.Response.StatusDescription
        }
        
        Write-TestResult "FAILED" "HTTP Status" "$StatusCode - $StatusDesc" -Color Red
    }
}

# Test some common variations
Write-Host "`n" -NoNewline
Write-TestResult "TEST" "Testing Common Variations" -Color Yellow

[Array]$Variations = @(
    "ClassicMapper.asmx/GetDriveMappings",
    "ClassicMapper.asmx/GetPrinterMappings", 
    "ClassicMapper.asmx/GetPSTMappings",
    "ClassicMapper.asmx/DriveMappings",
    "ClassicMapper.asmx/PrinterMappings",
    "ClassicMapper.asmx/PSTMappings",
    "ClassicMapper.asmx/GetMappings",
    "ClassicMapper.asmx/Mappings"
)

ForEach ($Variation in $Variations) {
    [String]$VariationURL = "$BaseURL/$Variation"
    Write-Host "`nTesting: $Variation" -ForegroundColor Cyan
    
    Try {
        [Object]$Response = Invoke-WebRequest -Uri $VariationURL -Method Get -UseBasicParsing -TimeoutSec 5
        Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
        Write-Host "  Content: $($Response.Content.Substring(0, [Math]::Min(100, $Response.Content.Length)))..." -ForegroundColor Gray
    } Catch {
        [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
        Write-TestResult "FAILED" "HTTP Status" $StatusCode -Color Red
    }
}

# Summary
Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "ENDPOINT TEST SUMMARY" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

Write-TestResult "INFO" "Endpoints Tested" $TestCount
Write-TestResult "INFO" "Endpoints Found" $FoundCount

If ($FoundCount -gt 0) {
    Write-Host "`nSUCCESS! Found $FoundCount working endpoint(s)." -ForegroundColor Green
    Write-Host "Check the WSDL files in C:\Temp\ for detailed method information." -ForegroundColor White
} Else {
    Write-Host "`nNo additional endpoints found." -ForegroundColor Red
    Write-Host "The mapping methods might be:" -ForegroundColor Yellow
    Write-Host "1. In a different server/domain" -ForegroundColor White
    Write-Host "2. Implemented as REST API instead of SOAP" -ForegroundColor White
    Write-Host "3. Removed or deprecated" -ForegroundColor White
    Write-Host "4. Require different authentication or access" -ForegroundColor White
}

Write-Host "`nNEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. Review any WSDL files created in C:\Temp\" -ForegroundColor White
Write-Host "2. Check with your backend team for the correct service URLs" -ForegroundColor White
Write-Host "3. Look for REST API endpoints instead of SOAP" -ForegroundColor White
Write-Host "4. Verify if the mapping functionality has been moved or changed" -ForegroundColor White

Write-Host "`n"
