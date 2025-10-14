<#
.SYNOPSIS
    Comprehensive WSDL and Method Name Testing Script
    
.DESCRIPTION
    Tests the WSDL structure and tries different method name variations
    to identify what methods actually exist on the backend server.
    
.EXAMPLE
    .\Test-WSDL-Methods.ps1
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
Write-Host "WSDL AND METHOD NAME TESTING" -ForegroundColor Magenta
Write-Host "========================================`n" -ForegroundColor Magenta

# Get current user info
[String]$Username = $env:USERNAME
[String]$ComputerName = $env:COMPUTERNAME

Write-TestResult "INFO" "Testing User" $Username
Write-TestResult "INFO" "Testing Computer" $ComputerName
Write-TestResult "INFO" "Server" "gdpmappercb.nomura.com"

# Test 1: Get and Analyze WSDL
Write-Host "`n" -NoNewline
Write-TestResult "TEST 1" "Analyzing WSDL Structure" -Color Yellow

Try {
    [Object]$WSDL = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx?WSDL" -UseBasicParsing
    Write-TestResult "SUCCESS" "WSDL Retrieved" "$($WSDL.Content.Length) bytes" -Color Green
    
    # Save WSDL to file for inspection
    [String]$WSDLPath = "C:\Temp\ClassicMapper.wsdl"
    $WSDL.Content | Out-File -FilePath $WSDLPath -Encoding UTF8
    Write-TestResult "INFO" "WSDL saved to" $WSDLPath -Color Cyan
    
    Write-Host "`n--- WSDL Content Analysis ---" -ForegroundColor Cyan
    
    # Search for various patterns
    [Array]$Patterns = @(
        @{ Name = "Operation Names"; Pattern = "operation.*name=" },
        @{ Name = "All Name Attributes"; Pattern = "name=" },
        @{ Name = "Method References"; Pattern = "method" },
        @{ Name = "Drive References"; Pattern = "Drive" },
        @{ Name = "Printer References"; Pattern = "Printer" },
        @{ Name = "PST References"; Pattern = "PST" },
        @{ Name = "Service Names"; Pattern = "service" },
        @{ Name = "Port Names"; Pattern = "port" },
        @{ Name = "Binding Names"; Pattern = "binding" }
    )
    
    ForEach ($PatternInfo in $Patterns) {
        Write-Host "`n$($PatternInfo.Name):" -ForegroundColor White
        [Array]$Matches = $WSDL.Content | Select-String -Pattern $PatternInfo.Pattern -AllMatches
        If ($Matches.Count -gt 0) {
            $Matches | ForEach-Object { 
                [String]$Line = $_.Line.Trim()
                If ($Line.Length -gt 0) {
                    Write-Host "  $Line" -ForegroundColor Gray
                }
            }
        } Else {
            Write-Host "  No matches found" -ForegroundColor DarkGray
        }
    }
    
    # Look for specific method patterns
    Write-Host "`n--- Specific Method Pattern Search ---" -ForegroundColor Cyan
    [Array]$MethodPatterns = @(
        "GetDrive",
        "GetPrinter", 
        "GetPST",
        "GetMapping",
        "GetMap",
        "ListDrive",
        "ListPrinter",
        "ListPST",
        "RetrieveDrive",
        "RetrievePrinter",
        "RetrievePST"
    )
    
    ForEach ($MethodPattern in $MethodPatterns) {
        [Array]$Matches = $WSDL.Content | Select-String -Pattern $MethodPattern -CaseSensitive:$False
        If ($Matches.Count -gt 0) {
            Write-Host "`nFound '$MethodPattern' references:" -ForegroundColor Green
            $Matches | ForEach-Object { 
                [String]$Line = $_.Line.Trim()
                If ($Line.Length -gt 0) {
                    Write-Host "  $Line" -ForegroundColor Gray
                }
            }
        }
    }
    
} Catch {
    Write-TestResult "ERROR" "Failed to retrieve WSDL" $_.Exception.Message -Color Red
    Exit 1
}

# Test 2: Try Different Method Name Variations
Write-Host "`n" -NoNewline
Write-TestResult "TEST 2" "Testing Method Name Variations" -Color Yellow

# Get DNs for testing
[String]$UserDN = "CN=chanwilw,OU=Users,OU=Resources,OU=HKG,OU=Nomura Wholesale,DC=RNDASIAPAC,DC=NOM"
[String]$ComputerDN = "CN=HKGWV030032,OU=VDI,OU=Devices,OU=Resources,OU=HKG,OU=Nomura Wholesale,DC=RNDASIAPAC,DC=NOM"

[Array]$TestMethods = @(
    # Drive mapping variations
    @{ Name = "GetDriveMappings"; Description = "Original (plural)" },
    @{ Name = "GetDriveMapping"; Description = "Singular" },
    @{ Name = "GetDriveMap"; Description = "Shortened" },
    @{ Name = "GetMappings"; Description = "Without Drive" },
    @{ Name = "GetDrives"; Description = "Just Drives" },
    @{ Name = "GetDriveList"; Description = "Alternative naming" },
    @{ Name = "ListDrives"; Description = "List prefix" },
    @{ Name = "RetrieveDrives"; Description = "Retrieve prefix" },
    @{ Name = "GetDriveMaps"; Description = "Maps plural" },
    @{ Name = "GetDriveMappingList"; Description = "Full descriptive" },
    
    # Printer mapping variations
    @{ Name = "GetPrinterMappings"; Description = "Original (plural)" },
    @{ Name = "GetPrinterMapping"; Description = "Singular" },
    @{ Name = "GetPrinters"; Description = "Just Printers" },
    @{ Name = "GetPrinterList"; Description = "List naming" },
    @{ Name = "ListPrinters"; Description = "List prefix" },
    
    # PST mapping variations
    @{ Name = "GetPSTMappings"; Description = "Original (plural)" },
    @{ Name = "GetPSTMapping"; Description = "Singular" },
    @{ Name = "GetPSTs"; Description = "Just PSTs" },
    @{ Name = "GetPSTList"; Description = "List naming" },
    @{ Name = "ListPSTs"; Description = "List prefix" },
    
    # Generic variations
    @{ Name = "GetMappings"; Description = "Generic mappings" },
    @{ Name = "GetAllMappings"; Description = "All mappings" },
    @{ Name = "GetUserMappings"; Description = "User mappings" },
    @{ Name = "GetComputerMappings"; Description = "Computer mappings" }
)

Write-Host "`nTesting $($TestMethods.Count) method name variations..." -ForegroundColor White

[Int]$SuccessCount = 0
[Int]$TestCount = 0

ForEach ($MethodInfo in $TestMethods) {
    $TestCount++
    [String]$MethodName = $MethodInfo.Name
    [String]$Description = $MethodInfo.Description
    
    Write-Host "`n[$TestCount] Testing: $MethodName ($Description)" -ForegroundColor Cyan
    
    # Build SOAP request
    [String]$TestSOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tem="http://tempuri.org/">
    <soap:Body>
        <tem:$MethodName>
            <tem:userDN>$UserDN</tem:userDN>
            <tem:computerDN>$ComputerDN</tem:computerDN>
        </tem:$MethodName>
    </soap:Body>
</soap:Envelope>
"@
    
    [Hashtable]$Headers = @{
        "Content-Type" = "text/xml; charset=utf-8"
        "SOAPAction" = "http://tempuri.org/$MethodName"
    }
    
    Try {
        [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $TestSOAP -Headers $Headers -TimeoutSec 5 -UseBasicParsing
        
        Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
        Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
        
        # Check if response contains success indicators
        If ($Response.Content -match "GetDriveMappingsResult|GetPrinterMappingsResult|GetPSTMappingsResult|Result|Success") {
            Write-TestResult "SUCCESS" "Response contains result indicators" -Color Green
        }
        
        # Show response preview
        [String]$Preview = $Response.Content.Substring(0, [Math]::Min(200, $Response.Content.Length))
        Write-Host "  Response Preview: $Preview..." -ForegroundColor Gray
        
        $SuccessCount++
        
    } Catch {
        [Int]$StatusCode = 0
        [String]$StatusDesc = "Unknown"
        
        If ($Null -ne $_.Exception.Response) {
            $StatusCode = [Int]$_.Exception.Response.StatusCode.value__
            $StatusDesc = $_.Exception.Response.StatusDescription
        }
        
        Write-TestResult "FAILED" "HTTP Status" "$StatusCode - $StatusDesc" -Color Red
        
        # Show error response if available
        Try {
            [Object]$ErrorStream = $_.Exception.Response.GetResponseStream()
            [Object]$Reader = New-Object System.IO.StreamReader($ErrorStream)
            [String]$ErrorBody = $Reader.ReadToEnd()
            
            If ($ErrorBody.Length -gt 0) {
                [String]$ErrorPreview = $ErrorBody.Substring(0, [Math]::Min(100, $ErrorBody.Length))
                Write-Host "  Error Preview: $ErrorPreview..." -ForegroundColor DarkRed
            }
        } Catch {
            # Could not read error body
        }
    }
}

# Test 3: Try Different Namespaces
Write-Host "`n" -NoNewline
Write-TestResult "TEST 3" "Testing Different Namespaces" -Color Yellow

[Array]$NamespaceTests = @(
    @{ Namespace = "http://tempuri.org/"; Prefix = "tem" },
    @{ Namespace = "http://schemas.microsoft.com/2003/10/Serialization/"; Prefix = "ser" },
    @{ Namespace = "http://schemas.datacontract.org/2004/07/"; Prefix = "dc" },
    @{ Namespace = "http://www.w3.org/2001/XMLSchema"; Prefix = "xsd" },
    @{ Namespace = "http://www.w3.org/2001/XMLSchema-instance"; Prefix = "xsi" },
    @{ Namespace = "http://nomura.com/"; Prefix = "nom" },
    @{ Namespace = "http://gdpmappercb.nomura.com/"; Prefix = "gdp" }
)

Write-Host "`nTesting different namespace combinations..." -ForegroundColor White

ForEach ($NSInfo in $NamespaceTests) {
    [String]$Namespace = $NSInfo.Namespace
    [String]$Prefix = $NSInfo.Prefix
    
    Write-Host "`nTesting namespace: $Namespace" -ForegroundColor Cyan
    
    [String]$NS_SOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:$Prefix="$Namespace">
    <soap:Body>
        <$Prefix`:GetDriveMappings>
            <$Prefix`:userDN>$UserDN</$Prefix`:userDN>
            <$Prefix`:computerDN>$ComputerDN</$Prefix`:computerDN>
        </$Prefix`:GetDriveMappings>
    </soap:Body>
</soap:Envelope>
"@
    
    [Hashtable]$NS_Headers = @{
        "Content-Type" = "text/xml; charset=utf-8"
        "SOAPAction" = "$Namespace/GetDriveMappings"
    }
    
    Try {
        [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $NS_SOAP -Headers $NS_Headers -TimeoutSec 5 -UseBasicParsing
        
        Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
        
    } Catch {
        [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
        Write-TestResult "FAILED" "HTTP Status" $StatusCode -Color Red
    }
}

# Summary
Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "TEST SUMMARY" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

Write-TestResult "INFO" "Methods Tested" $TestCount
Write-TestResult "INFO" "Successful Methods" $SuccessCount
Write-TestResult "INFO" "WSDL File Location" $WSDLPath

Write-Host "`nNEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. Review the WSDL file at: $WSDLPath" -ForegroundColor White
Write-Host "2. Look for any successful method names above" -ForegroundColor White
Write-Host "3. Check if any namespace combinations worked" -ForegroundColor White
Write-Host "4. If no methods worked, the service might be down or require different authentication" -ForegroundColor White
Write-Host "`n"

If ($SuccessCount -gt 0) {
    Write-Host "SUCCESS! Found working method(s). Update the PowerShell modules with the correct method names." -ForegroundColor Green
} Else {
    Write-Host "No working methods found. The service might require:" -ForegroundColor Red
    Write-Host "- Different authentication method" -ForegroundColor White
    Write-Host "- Different service endpoint" -ForegroundColor White
    Write-Host "- Service might be down or misconfigured" -ForegroundColor White
}

Write-Host "`n"
