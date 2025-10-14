<#
.SYNOPSIS
    Debug GetUserPSTs Method Issues
    
.DESCRIPTION
    Tests GetUserPSTs with different parameter combinations to identify the issue.
    
.EXAMPLE
    .\Test-PST-Debug.ps1
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
Write-Host "DEBUGGING GetUserPSTs METHOD" -ForegroundColor Magenta
Write-Host "========================================`n" -ForegroundColor Magenta

[String]$Username = $env:USERNAME
[String]$ComputerName = $env:COMPUTERNAME
[String]$UserDomain = "RNDASIAPAC"
[String]$UserOU = "Users"
[String]$UserSite = "TEST"
[String]$UserGroups = "<string>TestGroup1</string><string>TestGroup2</string>"

# Test 1: Try without HostName parameter
Write-Host "`n" -NoNewline
Write-TestResult "TEST 1" "GetUserPSTs WITHOUT HostName" -Color Yellow

[String]$PST_NoHost_SOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body>
  <GetUserPSTs xmlns="http://webtools.japan.nom">
   <UserId xsi:type="xsd:string">$Username</UserId>
   <Domain xsi:type="xsd:string">$UserDomain</Domain>
   <OuMapping xsi:type="xsd:string">$UserOU</OuMapping>
   <AdGroups xsi:type="xsd:string">$UserGroups</AdGroups>
   <Site xsi:type="xsd:string">$UserSite</Site>
  </GetUserPSTs>
 </soap:Body>
</soap:Envelope>
"@

[Hashtable]$PST_Headers = @{
    "Content-Type" = "text/xml; CharSet=UTF-8"
    "SOAPAction" = "http://webtools.japan.nom/GetUserPSTs"
}

Try {
    [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $PST_NoHost_SOAP -Headers $PST_Headers -TimeoutSec 10 -UseBasicParsing
    
    Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
    Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
    
    Write-Host "`n--- Response (No HostName) ---" -ForegroundColor Green
    Write-Host $Response.Content -ForegroundColor Gray
    
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
            Write-Host "`n--- Error Response (No HostName) ---" -ForegroundColor Red
            Write-Host $ErrorBody -ForegroundColor Gray
        }
    } Catch {
        Write-TestResult "INFO" "Could not read error response" -Color Yellow
    }
}

# Test 2: Try with HostName parameter
Write-Host "`n" -NoNewline
Write-TestResult "TEST 2" "GetUserPSTs WITH HostName" -Color Yellow

[String]$PST_WithHost_SOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body>
  <GetUserPSTs xmlns="http://webtools.japan.nom">
   <UserId xsi:type="xsd:string">$Username</UserId>
   <Domain xsi:type="xsd:string">$UserDomain</Domain>
   <OuMapping xsi:type="xsd:string">$UserOU</OuMapping>
   <AdGroups xsi:type="xsd:string">$UserGroups</AdGroups>
   <Site xsi:type="xsd:string">$UserSite</Site>
   <HostName xsi:type="xsd:string">$ComputerName</HostName>
  </GetUserPSTs>
 </soap:Body>
</soap:Envelope>
"@

Try {
    [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $PST_WithHost_SOAP -Headers $PST_Headers -TimeoutSec 10 -UseBasicParsing
    
    Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
    Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
    
    Write-Host "`n--- Response (With HostName) ---" -ForegroundColor Green
    Write-Host $Response.Content -ForegroundColor Gray
    
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
            Write-Host "`n--- Error Response (With HostName) ---" -ForegroundColor Red
            Write-Host $ErrorBody -ForegroundColor Gray
        }
    } Catch {
        Write-TestResult "INFO" "Could not read error response" -Color Yellow
    }
}

# Test 3: Try different method name variations
Write-Host "`n" -NoNewline
Write-TestResult "TEST 3" "Testing Method Name Variations" -Color Yellow

[Array]$PSTMethodVariations = @(
    "GetUserPSTs",
    "GetUserPST", 
    "GetPSTs",
    "GetPST",
    "GetUserPersonalFolders",
    "GetPersonalFolders",
    "GetUserPSTFiles",
    "GetPSTFiles"
)

ForEach ($MethodName in $PSTMethodVariations) {
    Write-Host "`nTesting method: $MethodName" -ForegroundColor Cyan
    
    [String]$TestSOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body>
  <$MethodName xmlns="http://webtools.japan.nom">
   <UserId xsi:type="xsd:string">$Username</UserId>
   <Domain xsi:type="xsd:string">$UserDomain</Domain>
   <OuMapping xsi:type="xsd:string">$UserOU</OuMapping>
   <AdGroups xsi:type="xsd:string">$UserGroups</AdGroups>
   <Site xsi:type="xsd:string">$UserSite</Site>
   <HostName xsi:type="xsd:string">$ComputerName</HostName>
  </$MethodName>
 </soap:Body>
</soap:Envelope>
"@
    
    [Hashtable]$TestHeaders = @{
        "Content-Type" = "text/xml; CharSet=UTF-8"
        "SOAPAction" = "http://webtools.japan.nom/$MethodName"
    }
    
    Try {
        [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $TestSOAP -Headers $TestHeaders -TimeoutSec 5 -UseBasicParsing
        
        Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
        Write-Host "  Response: $($Response.Content.Substring(0, [Math]::Min(100, $Response.Content.Length)))..." -ForegroundColor Gray
        
    } Catch {
        [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
        Write-TestResult "FAILED" "HTTP Status" $StatusCode -Color Red
    }
}

# Summary
Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "PST DEBUG SUMMARY" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

Write-Host "`nCURRENT STATUS:" -ForegroundColor Yellow
Write-Host "✅ GetUserDrives: WORKING" -ForegroundColor Green
Write-Host "✅ GetUserPrinters: WORKING" -ForegroundColor Green
Write-Host "❌ GetUserPSTs: FAILING" -ForegroundColor Red

Write-Host "`nPOSSIBLE ISSUES:" -ForegroundColor Yellow
Write-Host "1. Method name might be different (GetUserPST vs GetUserPSTs)" -ForegroundColor White
Write-Host "2. Missing required parameters" -ForegroundColor White
Write-Host "3. Different parameter structure" -ForegroundColor White
Write-Host "4. Service might not support PST functionality" -ForegroundColor White

Write-Host "`nNEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. Review the error responses above" -ForegroundColor White
Write-Host "2. Check if any method name variations worked" -ForegroundColor White
Write-Host "3. Update PowerShell modules with working methods" -ForegroundColor White
Write-Host "4. Handle PST functionality gracefully if not available" -ForegroundColor White

Write-Host "`n"
