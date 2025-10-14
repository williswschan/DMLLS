<#
.SYNOPSIS
    Test with Correct Method Names from VBScript
    
.DESCRIPTION
    Uses the EXACT method names and format from the working VBScript:
    - GetUserDrives (not GetDriveMappings)
    - GetUserPrinters (not GetPrinterMappings)  
    - GetUserPSTs (not GetPSTMappings)
    
.EXAMPLE
    .\Test-CorrectMethods.ps1
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
Write-Host "TESTING CORRECT METHOD NAMES FROM VBSCRIPT" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Green

# Get current user info
[String]$Username = $env:USERNAME
[String]$ComputerName = $env:COMPUTERNAME

Write-TestResult "INFO" "Testing User" $Username
Write-TestResult "INFO" "Testing Computer" $ComputerName
Write-TestResult "INFO" "Server" "gdpmappercb.nomura.com"
Write-TestResult "INFO" "Namespace" "http://webtools.japan.nom"

# Get DNs for testing
Try {
    Add-Type -AssemblyName System.DirectoryServices
    
    [Object]$UserSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $UserSearcher.Filter = "(&(objectClass=user)(samAccountName=$Username))"
    [Void]$UserSearcher.PropertiesToLoad.Add("distinguishedName")
    [Object]$UserResult = $UserSearcher.FindOne()
    
    [String]$UserDN = ""
    If ($UserResult) {
        $UserDN = $UserResult.Properties["distinguishedName"][0]
        Write-TestResult "SUCCESS" "User DN" $UserDN -Color Green
    } Else {
        Write-TestResult "WARNING" "User DN not found" "Using test values" -Color Yellow
        $UserDN = "CN=TestUser,OU=Users,DC=TEST,DC=COM"
    }
    
    [Object]$CompSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $CompSearcher.Filter = "(&(objectClass=computer)(name=$ComputerName))"
    [Void]$CompSearcher.PropertiesToLoad.Add("distinguishedName")
    [Object]$CompResult = $CompSearcher.FindOne()
    
    [String]$ComputerDN = ""
    If ($CompResult) {
        $ComputerDN = $CompResult.Properties["distinguishedName"][0]
        Write-TestResult "SUCCESS" "Computer DN" $ComputerDN -Color Green
    } Else {
        Write-TestResult "WARNING" "Computer DN not found" "Using test values" -Color Yellow
        $ComputerDN = "CN=TestPC,OU=Computers,DC=TEST,DC=COM"
    }
} Catch {
    Write-TestResult "ERROR" "Failed to get DNs" $_.Exception.Message -Color Red
    $UserDN = "CN=TestUser,OU=Users,DC=TEST,DC=COM"
    $ComputerDN = "CN=TestPC,OU=Computers,DC=TEST,DC=COM"
}

# Extract domain and OU info
[String]$UserDomain = "TEST"
[String]$UserOU = "OU=Users,DC=TEST,DC=COM"
[String]$UserSite = "TEST"
[String]$UserGroups = "<string>TestGroup1</string><string>TestGroup2</string>"

Try {
    If ($UserDN -match "DC=([^,]+)") {
        $UserDomain = $Matches[1]
    }
    If ($UserDN -match "OU=([^,]+)") {
        $UserOU = $Matches[1]
    }
} Catch {
    # Use defaults
}

# Test 1: GetUserDrives (the correct method name!)
Write-Host "`n" -NoNewline
Write-TestResult "TEST 1" "Testing GetUserDrives (CORRECT METHOD)" -Color Yellow

[String]$GetUserDrives_SOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body>
  <GetUserDrives xmlns="http://webtools.japan.nom">
   <UserId xsi:type="xsd:string">$Username</UserId>
   <Domain xsi:type="xsd:string">$UserDomain</Domain>
   <OuMapping xsi:type="xsd:string">$UserOU</OuMapping>
   <AdGroups xsi:type="xsd:string">$UserGroups</AdGroups>
   <Site xsi:type="xsd:string">$UserSite</Site>
  </GetUserDrives>
 </soap:Body>
</soap:Envelope>
"@

[Hashtable]$GetUserDrives_Headers = @{
    "Content-Type" = "text/xml; CharSet=UTF-8"
    "SOAPAction" = "http://webtools.japan.nom/GetUserDrives"
}

Write-Host "SOAP Request:" -ForegroundColor White
Write-Host $GetUserDrives_SOAP -ForegroundColor Gray

Try {
    [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $GetUserDrives_SOAP -Headers $GetUserDrives_Headers -TimeoutSec 10 -UseBasicParsing
    
    Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
    Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
    
    Write-Host "`n--- GetUserDrives Response ---" -ForegroundColor Green
    Write-Host $Response.Content -ForegroundColor Gray
    
    # Parse response
    Try {
        [Xml]$ResponseXML = $Response.Content
        [Object]$GetUserDrivesResult = $ResponseXML.SelectSingleNode("//*[local-name()='GetUserDrivesResult']")
        
        If ($GetUserDrivesResult) {
            Write-TestResult "SUCCESS" "GetUserDrives Result found" -Color Green
            
            # Look for MapperDrive objects
            [Array]$MapperDrives = $ResponseXML.SelectNodes("//*[local-name()='MapperDrive']")
            Write-TestResult "INFO" "MapperDrive objects found" $MapperDrives.Count -Color Cyan
            
            If ($MapperDrives.Count -gt 0) {
                Write-Host "`n--- Drive Details ---" -ForegroundColor Cyan
                ForEach ($Drive in $MapperDrives) {
                    [Object]$Id = $Drive.SelectSingleNode("*[local-name()='Id']")
                    [Object]$DriveLetter = $Drive.SelectSingleNode("*[local-name()='Drive']")
                    [Object]$UncPath = $Drive.SelectSingleNode("*[local-name()='UncPath']")
                    [Object]$Description = $Drive.SelectSingleNode("*[local-name()='Description']")
                    
                    Write-Host "  Drive:" -ForegroundColor White
                    If ($Id) { Write-Host "    Id: $($Id.InnerText)" -ForegroundColor Gray }
                    If ($DriveLetter) { Write-Host "    Drive: $($DriveLetter.InnerText)" -ForegroundColor Gray }
                    If ($UncPath) { Write-Host "    UncPath: $($UncPath.InnerText)" -ForegroundColor Gray }
                    If ($Description) { Write-Host "    Description: $($Description.InnerText)" -ForegroundColor Gray }
                    Write-Host ""
                }
            }
        } Else {
            Write-TestResult "WARNING" "Could not parse GetUserDrivesResult" -Color Yellow
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

# Test 2: GetUserPrinters
Write-Host "`n" -NoNewline
Write-TestResult "TEST 2" "Testing GetUserPrinters" -Color Yellow

[String]$GetUserPrinters_SOAP = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
 <soap:Body>
  <GetUserPrinters xmlns="http://webtools.japan.nom">
   <UserId xsi:type="xsd:string">$Username</UserId>
   <Domain xsi:type="xsd:string">$UserDomain</Domain>
   <OuMapping xsi:type="xsd:string">$UserOU</OuMapping>
   <AdGroups xsi:type="xsd:string">$UserGroups</AdGroups>
   <Site xsi:type="xsd:string">$UserSite</Site>
   <HostName xsi:type="xsd:string">$ComputerName</HostName>
  </GetUserPrinters>
 </soap:Body>
</soap:Envelope>
"@

[Hashtable]$GetUserPrinters_Headers = @{
    "Content-Type" = "text/xml; CharSet=UTF-8"
    "SOAPAction" = "http://webtools.japan.nom/GetUserPrinters"
}

Try {
    [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $GetUserPrinters_SOAP -Headers $GetUserPrinters_Headers -TimeoutSec 10 -UseBasicParsing
    
    Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
    Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
    
    Write-Host "`n--- GetUserPrinters Response ---" -ForegroundColor Green
    Write-Host $Response.Content -ForegroundColor Gray
    
} Catch {
    [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
    [String]$StatusDesc = if ($_.Exception.Response) { $_.Exception.Response.StatusDescription } else { "Unknown" }
    
    Write-TestResult "FAILED" "HTTP Status" "$StatusCode - $StatusDesc" -Color Red
}

# Test 3: GetUserPSTs
Write-Host "`n" -NoNewline
Write-TestResult "TEST 3" "Testing GetUserPSTs" -Color Yellow

[String]$GetUserPSTs_SOAP = @"
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

[Hashtable]$GetUserPSTs_Headers = @{
    "Content-Type" = "text/xml; CharSet=UTF-8"
    "SOAPAction" = "http://webtools.japan.nom/GetUserPSTs"
}

Try {
    [Object]$Response = Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method Post -Body $GetUserPSTs_SOAP -Headers $GetUserPSTs_Headers -TimeoutSec 10 -UseBasicParsing
    
    Write-TestResult "SUCCESS" "HTTP Status" $Response.StatusCode -Color Green
    Write-TestResult "SUCCESS" "Content Length" "$($Response.Content.Length) bytes" -Color Green
    
    Write-Host "`n--- GetUserPSTs Response ---" -ForegroundColor Green
    Write-Host $Response.Content -ForegroundColor Gray
    
} Catch {
    [Int]$StatusCode = if ($_.Exception.Response) { $_.Exception.Response.StatusCode.value__ } else { 0 }
    [String]$StatusDesc = if ($_.Exception.Response) { $_.Exception.Response.StatusDescription } else { "Unknown" }
    
    Write-TestResult "FAILED" "HTTP Status" "$StatusCode - $StatusDesc" -Color Red
}

# Summary
Write-Host "`n========================================" -ForegroundColor Magenta
Write-Host "CORRECT METHOD TEST SUMMARY" -ForegroundColor Magenta
Write-Host "========================================" -ForegroundColor Magenta

Write-Host "`nKEY FINDINGS:" -ForegroundColor Yellow
Write-Host "1. VBScript uses GetUserDrives (not GetDriveMappings)" -ForegroundColor White
Write-Host "2. VBScript uses GetUserPrinters (not GetPrinterMappings)" -ForegroundColor White
Write-Host "3. VBScript uses GetUserPSTs (not GetPSTMappings)" -ForegroundColor White
Write-Host "4. VBScript uses namespace http://webtools.japan.nom" -ForegroundColor White
Write-Host "5. VBScript uses different parameter structure" -ForegroundColor White

Write-Host "`nNEXT STEPS:" -ForegroundColor Yellow
Write-Host "1. If any tests succeeded, update PowerShell modules with correct method names" -ForegroundColor White
Write-Host "2. Update SOAP request format to match VBScript exactly" -ForegroundColor White
Write-Host "3. Update parameter names and structure" -ForegroundColor White
Write-Host "4. Test with real user data" -ForegroundColor White

Write-Host "`n"
