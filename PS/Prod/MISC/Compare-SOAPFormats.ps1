<#
.SYNOPSIS
    Compares different SOAP request formats side-by-side
    
.DESCRIPTION
    Shows the exact differences between various SOAP formats to identify
    what the backend server expects vs what we're sending.
    
.EXAMPLE
    .\Compare-SOAPFormats.ps1
#>

[CmdletBinding()]
Param()

[String]$Username = $env:USERNAME
[String]$UserDN = "CN=Test User,OU=Users,DC=DOMAIN,DC=COM"
[String]$ComputerDN = "CN=TestPC,OU=Computers,DC=DOMAIN,DC=COM"

Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "SOAP FORMAT COMPARISON" -ForegroundColor Magenta
Write-Host "========================================`n" -ForegroundColor Magenta

# Format 1: Current Implementation (Both HTTP Basic + SOAP Header Auth)
Write-Host "FORMAT 1: Current Implementation (HTTP Basic + SOAP Header)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

[String]$Format1_Body = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tem="http://tempuri.org/">
    <soap:Header>
        <tem:AuthHeader>
            <tem:Username>$Username</tem:Username>
            <tem:Password>placeholder</tem:Password>
        </tem:AuthHeader>
    </soap:Header>
    <soap:Body>
        <tem:GetDriveMappings>
            <tem:userDN>$UserDN</tem:userDN>
            <tem:computerDN>$ComputerDN</tem:computerDN>
        </tem:GetDriveMappings>
    </soap:Body>
</soap:Envelope>
"@

[Hashtable]$Format1_Headers = @{
    "Content-Type" = "text/xml; charset=utf-8"
    "SOAPAction" = "http://tempuri.org/GetDriveMappings"
    "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$Username`:placeholder"))
}

Write-Host "Headers:" -ForegroundColor White
$Format1_Headers.GetEnumerator() | ForEach-Object {
    If ($_.Key -eq "Authorization") {
        Write-Host "  $($_.Key): Basic [Base64 of '${Username}:placeholder']" -ForegroundColor Gray
    } Else {
        Write-Host "  $($_.Key): $($_.Value)" -ForegroundColor Gray
    }
}

Write-Host "`nSOAP Body:" -ForegroundColor White
Write-Host $Format1_Body -ForegroundColor Gray

# Format 2: HTTP Basic Auth Only (No SOAP Header)
Write-Host "`n`nFORMAT 2: HTTP Basic Auth Only" -ForegroundColor Cyan
Write-Host "===============================" -ForegroundColor Cyan

[String]$Format2_Body = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" xmlns:tem="http://tempuri.org/">
    <soap:Body>
        <tem:GetDriveMappings>
            <tem:userDN>$UserDN</tem:userDN>
            <tem:computerDN>$ComputerDN</tem:computerDN>
        </tem:GetDriveMappings>
    </soap:Body>
</soap:Envelope>
"@

[Hashtable]$Format2_Headers = @{
    "Content-Type" = "text/xml; charset=utf-8"
    "SOAPAction" = "http://tempuri.org/GetDriveMappings"
    "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes("$Username`:placeholder"))
}

Write-Host "Headers:" -ForegroundColor White
$Format2_Headers.GetEnumerator() | ForEach-Object {
    If ($_.Key -eq "Authorization") {
        Write-Host "  $($_.Key): Basic [Base64 of '${Username}:placeholder']" -ForegroundColor Gray
    } Else {
        Write-Host "  $($_.Key): $($_.Value)" -ForegroundColor Gray
    }
}

Write-Host "`nSOAP Body:" -ForegroundColor White
Write-Host $Format2_Body -ForegroundColor Gray

# Format 3: Windows Authentication
Write-Host "`n`nFORMAT 3: Windows Authentication (UseDefaultCredentials)" -ForegroundColor Cyan
Write-Host "=========================================================" -ForegroundColor Cyan

[Hashtable]$Format3_Headers = @{
    "Content-Type" = "text/xml; charset=utf-8"
    "SOAPAction" = "http://tempuri.org/GetDriveMappings"
}

Write-Host "Headers:" -ForegroundColor White
$Format3_Headers.GetEnumerator() | ForEach-Object {
    Write-Host "  $($_.Key): $($_.Value)" -ForegroundColor Gray
}
Write-Host "  UseDefaultCredentials: True" -ForegroundColor Gray

Write-Host "`nSOAP Body: (Same as Format 2)" -ForegroundColor White

# Format 4: User-Provided Working Format (From earlier conversation)
Write-Host "`n`nFORMAT 4: User-Provided Working Code Format" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan

Write-Host @"
This is the format from the working PowerShell code you provided earlier.
It should be IDENTICAL to Format 1 (our current implementation).

If the working code uses a different:
- Namespace (e.g., not 'tem' or 'http://tempuri.org/')
- Element names (e.g., different parameter names)
- XML structure
Please note the differences!
"@ -ForegroundColor Yellow

# Key Differences to Check
Write-Host "`n`nKEY DIFFERENCES TO CHECK:" -ForegroundColor Magenta
Write-Host "=========================" -ForegroundColor Magenta
Write-Host "`n1. Namespace Prefix:" -ForegroundColor White
Write-Host "   - Current: xmlns:tem='http://tempuri.org/'" -ForegroundColor Gray
Write-Host "   - Check if should be different (e.g., xmlns:web, xmlns:ns1, etc.)" -ForegroundColor Gray

Write-Host "`n2. SOAP Action:" -ForegroundColor White
Write-Host "   - Current: http://tempuri.org/GetDriveMappings" -ForegroundColor Gray
Write-Host "   - Check if should be different namespace" -ForegroundColor Gray

Write-Host "`n3. Parameter Names:" -ForegroundColor White
Write-Host "   - Current: <tem:userDN>, <tem:computerDN>" -ForegroundColor Gray
Write-Host "   - Check case sensitivity (userDN vs userDn vs UserDN)" -ForegroundColor Gray

Write-Host "`n4. Authentication:" -ForegroundColor White
Write-Host "   - Current: Both HTTP Basic + SOAP Header" -ForegroundColor Gray
Write-Host "   - Check if both are required or just one" -ForegroundColor Gray

Write-Host "`n5. Content-Type:" -ForegroundColor White
Write-Host "   - Current: text/xml; charset=utf-8" -ForegroundColor Gray
Write-Host "   - Some services require: application/soap+xml" -ForegroundColor Gray

Write-Host "`n"
Write-Host "NEXT STEP: Run Test-WebService.ps1 to see actual server responses" -ForegroundColor Yellow
Write-Host "`n"

