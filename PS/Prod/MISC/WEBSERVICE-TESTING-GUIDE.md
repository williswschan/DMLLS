# Web Service Testing Guide

## Overview
The DMLLS PowerShell suite currently returns **400 Bad Request** when calling the backend web services. These diagnostic scripts will help identify and fix the issue.

## Available Diagnostic Scripts

### 1. `Test-MockBackend.ps1` - Start Here
**Purpose:** Verify basic connectivity and identify if you're hitting the mock or real server.

**What it tests:**
- DNS resolution for `gdpmappercb.nomura.com`
- HTTP connectivity to the base URL
- Service endpoint availability (ClassicMapper.asmx, ClassicInventory.asmx)
- WSDL availability and method listing

**Run:**
```powershell
.\MISC\Test-MockBackend.ps1
```

**Expected Output:**
- DNS resolves to IP address
- HTTP 200 on base URL
- Service endpoints accessible
- Lists available SOAP methods

---

### 2. `Compare-SOAPFormats.ps1` - Review Formats
**Purpose:** Compare different SOAP request formats side-by-side.

**What it shows:**
- Format 1: Current implementation (HTTP Basic + SOAP Header Auth)
- Format 2: HTTP Basic Auth only
- Format 3: Windows Authentication
- Format 4: Reference to your working code
- Key differences to check (namespaces, parameter names, etc.)

**Run:**
```powershell
.\MISC\Compare-SOAPFormats.ps1
```

---

### 3. `Test-WebService.ps1` - Basic Testing
**Purpose:** Test actual SOAP requests with detailed logging.

**What it tests:**
- Auto-detects User and Computer DNs
- Sends GetDriveMappings request to Mapper service
- Sends InsertLogonInventory request to Inventory service
- Tests alternative format (no SOAP header auth)
- Tests Windows authentication

**Run:**
```powershell
# Test both services
.\MISC\Test-WebService.ps1

# Test only Mapper
.\MISC\Test-WebService.ps1 -ServiceType Mapper

# Test only Inventory
.\MISC\Test-WebService.ps1 -ServiceType Inventory

# Specify DNs manually
.\MISC\Test-WebService.ps1 -UserDN "CN=Willis Chan,OU=Individual,OU=Users,OU=HK,DC=MYMSNGROUP,DC=COM" -ComputerDN "CN=HKNOM01,OU=NOM,DC=MYMSNGROUP,DC=COM"
```

**Expected Output:**
- Shows complete SOAP request
- Shows all HTTP headers
- Shows HTTP status code
- Shows complete SOAP response (formatted XML)
- Shows parsed results
- Shows detailed error messages for failures

---

### 4. `Test-WebService-Advanced.ps1` - Authentication Matrix
**Purpose:** Test 5 different authentication combinations to find what works.

**What it tests:**
1. HTTP Basic + SOAP Header (current method)
2. HTTP Basic only
3. SOAP Header only
4. Windows Authentication
5. No authentication

**Run:**
```powershell
.\MISC\Test-WebService-Advanced.ps1
```

**Expected Output:**
- Tests each combination
- Shows which ones succeed (HTTP 200)
- Shows response previews
- Identifies working authentication method

---

### 5. `Test-WorkingSOAPCode.ps1` - Working Code Validation
**Purpose:** Use the exact SOAP format from your working PowerShell code.

**What it does:**
- Uses the format you confirmed works in your environment
- Compares response with our implementation
- Shows XML namespaces and structure
- Tests multiple XPath queries to find correct parsing method

**Run:**
```powershell
.\MISC\Test-WorkingSOAPCode.ps1
```

**Expected Output:**
- If working code succeeds: Shows response structure and XPath matches
- If working code fails: Confirms network/DNS issues

---

## Testing Strategy

### Phase 1: Verify Connectivity
```powershell
.\MISC\Test-MockBackend.ps1
```
**Goal:** Ensure server is reachable and endpoints exist

### Phase 2: Test Working Code
```powershell
.\MISC\Test-WorkingSOAPCode.ps1
```
**Goal:** Verify the format you provided actually works

### Phase 3: Test Authentication Methods
```powershell
.\MISC\Test-WebService-Advanced.ps1
```
**Goal:** Find which authentication method the server accepts

### Phase 4: Test Full Request
```powershell
.\MISC\Test-WebService.ps1
```
**Goal:** Test complete request with real DNs and parse response

### Phase 5: Compare Formats
```powershell
.\MISC\Compare-SOAPFormats.ps1
```
**Goal:** Review differences and identify issues

---

## Common Issues and Solutions

### Issue 1: 400 Bad Request
**Possible Causes:**
- Incorrect SOAP namespace
- Wrong parameter names (case sensitivity)
- Malformed XML
- Missing required fields

**How to diagnose:**
- Check error response body for SOAP fault details
- Compare request with WSDL definition
- Verify parameter names match exactly

### Issue 2: 401 Unauthorized
**Possible Causes:**
- Wrong authentication method
- Invalid credentials
- Missing Authorization header

**How to diagnose:**
- Run Test-WebService-Advanced.ps1 to test all auth methods
- Check if server requires Windows authentication

### Issue 3: 500 Internal Server Error
**Possible Causes:**
- Server-side processing error
- Database connection issue
- Invalid DN format
- Backend service crash

**How to diagnose:**
- Check error response for SOAP fault
- Verify DNs are valid and properly formatted
- Check server logs (if accessible)

### Issue 4: 404 Not Found
**Possible Causes:**
- Wrong service URL
- Service not deployed
- IIS routing issue

**How to diagnose:**
- Verify URL with: http://gdpmappercb.nomura.com/ClassicMapper.asmx
- Check if WSDL is accessible

---

## After Running Diagnostics

Once you've identified the issue using these scripts:

1. **If authentication method needs changing:**
   - Update `Send-DMSOAPRequestWithAuth` in `DMServiceCommon.psm1`
   - Modify header construction or add `UseDefaultCredentials`

2. **If SOAP format needs changing:**
   - Update SOAP body templates in `DMMapperService.psm1` and `DMInventoryService.psm1`
   - Adjust namespace prefixes or parameter names

3. **If response parsing needs fixing:**
   - Update XPath queries in mapper/inventory modules
   - Adjust SelectNodes() calls to match actual response structure

4. **If it's a backend issue:**
   - Work with backend team to fix server configuration
   - Or update mock backend to match real server behavior

---

## Quick Reference

### Current Implementation Location
- **Service Common:** `Modules\Services\DMServiceCommon.psm1`
- **Mapper Service:** `Modules\Services\DMMapperService.psm1`
- **Inventory Service:** `Modules\Services\DMInventoryService.psm1`

### Current Format Being Used
- **Authentication:** HTTP Basic + SOAP Header AuthHeader
- **Namespace:** `xmlns:tem="http://tempuri.org/"`
- **SOAP Action:** `http://tempuri.org/GetDriveMappings`
- **Parameters:** `<tem:userDN>`, `<tem:computerDN>`

---

## Need More Help?

Run all tests and save output to a file for review:
```powershell
.\MISC\Test-MockBackend.ps1 > C:\Temp\webservice-test-results.txt 2>&1
.\MISC\Test-WebService-Advanced.ps1 >> C:\Temp\webservice-test-results.txt 2>&1
.\MISC\Test-WorkingSOAPCode.ps1 >> C:\Temp\webservice-test-results.txt 2>&1
.\MISC\Test-WebService.ps1 >> C:\Temp\webservice-test-results.txt 2>&1

# Review the results
notepad C:\Temp\webservice-test-results.txt
```

