# Desktop Management Mock Backend - Test Script
# Tests all SOAP endpoints to verify server is working

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Desktop Management Backend Test Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

[String]$BaseUrl = "http://gdpmappercb.nomura.com"
[Int]$TestsPassed = 0
[Int]$TestsFailed = 0

Function Test-Endpoint {
    Param(
        [String]$ServiceName,
        [String]$MethodName,
        [String]$SoapBody
    )
    
    Write-Host "Testing: $ServiceName -> $MethodName ... " -NoNewline
    
    [String]$Url = "$BaseUrl/$ServiceName"
    [String]$FullSoapRequest = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Body>
        $SoapBody
    </soap:Body>
</soap:Envelope>
"@
    
    Try {
        [Object]$Response = Invoke-WebRequest -Uri $Url -Method POST -Body $FullSoapRequest -ContentType "text/xml" -ErrorAction Stop
        
        If ($Response.StatusCode -eq 200) {
            Write-Host "[OK]" -ForegroundColor Green
            $Script:TestsPassed++
            Return $True
        } Else {
            Write-Host "[FAILED] Status: $($Response.StatusCode)" -ForegroundColor Red
            $Script:TestsFailed++
            Return $False
        }
    } Catch {
        Write-Host "[FAILED] Error: $($_.Exception.Message)" -ForegroundColor Red
        $Script:TestsFailed++
        Return $False
    }
}

# Test DNS Resolution
Write-Host "Checking DNS Resolution..." -ForegroundColor Yellow
Try {
    [Object]$DnsResult = Resolve-DnsName -Name "gdpmappercb.nomura.com" -ErrorAction Stop
    [String]$ResolvedIP = $DnsResult[0].IPAddress
    Write-Host "  gdpmappercb.nomura.com -> $ResolvedIP" -ForegroundColor Green
    
    If ($ResolvedIP -ne "127.0.0.1") {
        Write-Host "  [WARNING] Not resolving to localhost (127.0.0.1)" -ForegroundColor Yellow
        Write-Host "  This may be intentional if running on a different server" -ForegroundColor Yellow
    }
} Catch {
    Write-Host "  [ERROR] Cannot resolve gdpmappercb.nomura.com" -ForegroundColor Red
    Write-Host "  Run setup_hosts.bat as Administrator to configure hosts file" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Running Endpoint Tests..." -ForegroundColor Yellow
Write-Host ""

# Test Mapper Service
Write-Host "=== Mapper Service (ClassicMapper.asmx) ===" -ForegroundColor Cyan

Test-Endpoint -ServiceName "ClassicMapper.asmx" -MethodName "TestService" -SoapBody @"
<TestService xmlns="http://webtools.japan.nom" />
"@

Test-Endpoint -ServiceName "ClassicMapper.asmx" -MethodName "GetUserDrives" -SoapBody @"
<GetUserDrives xmlns="http://webtools.japan.nom">
    <UserId>testuser</UserId>
    <Domain>ASIAPAC.NOM</Domain>
    <OuMapping>RESOURCES/HKG/USERS</OuMapping>
    <AdGroups>Domain Users</AdGroups>
    <Site>HKG</Site>
</GetUserDrives>
"@

Test-Endpoint -ServiceName "ClassicMapper.asmx" -MethodName "GetUserPrinters" -SoapBody @"
<GetUserPrinters xmlns="http://webtools.japan.nom">
    <HostName>WKS001</HostName>
    <Domain>ASIAPAC.NOM</Domain>
    <AdGroups>Domain Computers</AdGroups>
    <Site>HKG</Site>
</GetUserPrinters>
"@

Test-Endpoint -ServiceName "ClassicMapper.asmx" -MethodName "GetUserPersonalFolders" -SoapBody @"
<GetUserPersonalFolders xmlns="http://webtools.japan.nom">
    <UserId>testuser</UserId>
    <Domain>ASIAPAC.NOM</Domain>
    <OuMapping>RESOURCES/HKG/USERS</OuMapping>
</GetUserPersonalFolders>
"@

Write-Host ""

# Test Inventory Service
Write-Host "=== Inventory Service (ClassicInventory.asmx) ===" -ForegroundColor Cyan

Test-Endpoint -ServiceName "ClassicInventory.asmx" -MethodName "TestService" -SoapBody @"
<TestService xmlns="http://webtools.japan.nom" />
"@

Test-Endpoint -ServiceName "ClassicInventory.asmx" -MethodName "InsertLogonInventory" -SoapBody @"
<InsertLogonInventory xmlns="http://webtools.japan.nom">
    <UserId>testuser</UserId>
    <UserDomain>ASIAPAC.NOM</UserDomain>
    <HostName>WKS001</HostName>
    <Domain>ASIAPAC.NOM</Domain>
    <SiteName>HKG</SiteName>
    <City>HKG</City>
    <OuMapping>RESOURCES/HKG/DEVICES</OuMapping>
</InsertLogonInventory>
"@

Test-Endpoint -ServiceName "ClassicInventory.asmx" -MethodName "InsertActiveDriveMappingsFromInventory" -SoapBody @"
<InsertActiveDriveMappingsFromInventory xmlns="http://webtools.japan.nom">
    <UserId>testuser</UserId>
    <HostName>WKS001</HostName>
    <Domain>ASIAPAC.NOM</Domain>
    <SiteName>HKG</SiteName>
    <City>HKG</City>
    <Drive>H:</Drive>
    <UncPath>\\fileserver\home\testuser</UncPath>
    <Description>Home Drive</Description>
    <OuMapping>RESOURCES/HKG/DEVICES</OuMapping>
</InsertActiveDriveMappingsFromInventory>
"@

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Test Results" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Passed: $TestsPassed" -ForegroundColor Green
Write-Host "Failed: $TestsFailed" -ForegroundColor $(If ($TestsFailed -gt 0) { "Red" } Else { "Green" })
Write-Host ""

If ($TestsFailed -eq 0) {
    Write-Host "[SUCCESS] All tests passed! Backend is working correctly." -ForegroundColor Green
} Else {
    Write-Host "[WARNING] Some tests failed. Check server logs for details." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "To view collected data, check CSV files in: WEB\Data\" -ForegroundColor Cyan
Write-Host ""

