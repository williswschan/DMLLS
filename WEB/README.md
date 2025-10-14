# Desktop Management Mock Backend Server

A lightweight mock backend server that simulates the production IIS SOAP web services for Desktop Management Suite testing.

## Overview

This Python Flask application mimics:
- `gdpmappercb.nomura.com/ClassicMapper.asmx` (Mapper Service)
- `gdpmappercb.nomura.com/ClassicInventory.asmx` (Inventory Service)

**Storage:** CSV files (no database required)  
**Protocol:** SOAP over HTTP  
**Port:** 80 (HTTP)

## Features

### Mapper Service (ClassicMapper.asmx)

✅ **TestService** - Health check  
✅ **GetUserDrives** - Returns drive mappings from `drives.csv`  
✅ **GetUserPrinters** - Returns printer mappings from `printers.csv`  
✅ **GetUserPersonalFolders** - Returns PST mappings from `pst.csv`

### Inventory Service (ClassicInventory.asmx)

✅ **TestService** - Health check  
✅ **InsertLogonInventory** - Saves session logon to `sessions.csv`  
✅ **InsertLogoffInventory** - Saves session logoff to `sessions.csv`  
✅ **InsertActiveDriveMappingsFromInventory** - Saves to `inventory_drives.csv`  
✅ **InsertMapperPrinterInventory** - Saves to `inventory_printers.csv`  
✅ **InsertActivePersonalFolderMappingsFromInventory** - Saves to `inventory_pst.csv`

## Quick Start

### Prerequisites

1. **Python 3.7+** installed
2. **Administrator privileges** (required for port 80)

### Installation

```bash
# Navigate to WEB directory
cd C:\Temp\Script\DMLLS\WEB

# Install Flask
pip install -r requirements.txt
```

### Run the Server

```bash
# Start the mock backend (requires admin privileges for port 80)
python mock_backend.py
```

Server will start on: `http://0.0.0.0:80`

### Configure DNS Redirection

**Option 1: Windows Hosts File (Recommended for Testing)**

1. Open `C:\Windows\System32\drivers\etc\hosts` as Administrator
2. Add this line:
   ```
   127.0.0.1    gdpmappercb.nomura.com
   ```
3. Save and close

**Option 2: DNS Server**

Configure your DNS server to resolve `gdpmappercb.nomura.com` to the IP address where this mock server is running.

### Verify It's Working

Open browser and navigate to:
- `http://gdpmappercb.nomura.com/` - Should show status page
- `http://gdpmappercb.nomura.com/ClassicMapper.asmx` - Mapper service
- `http://gdpmappercb.nomura.com/ClassicInventory.asmx` - Inventory service

## Data Storage

### Mapper Data (Source - what to map TO clients)

**`Data/drives.csv`** - Drive mappings to send to clients
```csv
Id,Domain,UserId,AdGroup,Site,Drive,UncPath,Description,DisconnectOnLogin
1,ASIAPAC.NOM,testuser,Domain Users,HKG,H:,\\fileserver\home\testuser,Home Drive,false
```

**`Data/printers.csv`** - Printer mappings to send to clients
```csv
Id,HostName,UncPath,IsDefault,Description
1,WKS001,\\printserver\printer01,true,Office Printer 1
```

**`Data/pst.csv`** - PST mappings to send to clients
```csv
Id,UserId,UncPath,DisconnectOnLogin
1,testuser,\\fileserver\pst\testuser\archive.pst,false
```

### Inventory Data (Destination - what clients SEND to backend)

**`Data/sessions.csv`** - Logon/Logoff events  
**`Data/inventory_drives.csv`** - Drive inventory from clients  
**`Data/inventory_printers.csv`** - Printer inventory from clients  
**`Data/inventory_pst.csv`** - PST inventory from clients

These files are auto-created when clients send data.

## Testing with PowerShell

### Test Mapper Service

```powershell
# Get drive mappings for user
Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicMapper.asmx" -Method POST -Body @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Body>
        <GetUserDrives xmlns="http://webtools.japan.nom">
            <UserId>testuser</UserId>
            <Domain>ASIAPAC.NOM</Domain>
            <OuMapping>RESOURCES/HKG/USERS</OuMapping>
        </GetUserDrives>
    </soap:Body>
</soap:Envelope>
"@ -ContentType "text/xml"
```

### Test Inventory Service

```powershell
# Send logon inventory
Invoke-WebRequest -Uri "http://gdpmappercb.nomura.com/ClassicInventory.asmx" -Method POST -Body @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Body>
        <InsertLogonInventory xmlns="http://webtools.japan.nom">
            <UserId>testuser</UserId>
            <UserDomain>ASIAPAC.NOM</UserDomain>
            <HostName>WKS001</HostName>
            <Domain>ASIAPAC.NOM</Domain>
            <SiteName>HKG</SiteName>
            <City>HKG</City>
            <OuMapping>RESOURCES/HKG/DEVICES</OuMapping>
        </InsertLogonInventory>
    </soap:Body>
</soap:Envelope>
"@ -ContentType "text/xml"
```

## Customizing Test Data

### Add Test Users/Computers

Edit the CSV files in `Data/` directory:

**For Drives:**
- Edit `drives.csv`
- Add rows with `UserId` matching your test user
- Server filters by `UserId` when responding to `GetUserDrives`

**For Printers:**
- Edit `printers.csv`
- Add rows with `HostName` matching your test computer
- Server filters by `HostName` when responding to `GetUserPrinters`

**For PST Files:**
- Edit `pst.csv`
- Add rows with `UserId` matching your test user
- Server filters by `UserId` when responding to `GetUserPersonalFolders`

### View Collected Inventory

Check the auto-created CSV files:
- `Data/sessions.csv` - Logon/logoff events
- `Data/inventory_drives.csv` - Drives reported by clients
- `Data/inventory_printers.csv` - Printers reported by clients
- `Data/inventory_pst.csv` - PST files reported by clients

## Troubleshooting

### Port 80 Access Denied

**Problem:** Cannot bind to port 80 without admin rights

**Solution 1:** Run PowerShell/CMD as Administrator, then run `python mock_backend.py`

**Solution 2:** Change port in `mock_backend.py`:
```python
# Change line at bottom of file:
app.run(host='0.0.0.0', port=8080, debug=True)  # Use 8080 instead of 80
```

Then update hosts file:
```
127.0.0.1:8080    gdpmappercb.nomura.com
```

### DNS Not Resolving

**Test DNS:**
```powershell
nslookup gdpmappercb.nomura.com
# Should return 127.0.0.1 or your server IP
```

**Clear DNS Cache:**
```powershell
ipconfig /flushdns
```

### Server Not Responding

**Check if server is running:**
```powershell
netstat -ano | findstr :80
# Should show LISTENING on port 80
```

**Check Windows Firewall:**
```powershell
# Allow Python through firewall
netsh advfirewall firewall add rule name="Python Mock Backend" dir=in action=allow program="C:\Path\To\python.exe" enable=yes
```

## Logs

Server logs all incoming requests to console:
```
[2025-10-13 17:30:45] ClassicMapper -> GetUserDrives
  UserId: testuser
  Domain: ASIAPAC.NOM
  OuMapping: RESOURCES/HKG/USERS
```

## Architecture

```
PowerShell Client
    ↓ SOAP Request
    ↓ (http://gdpmappercb.nomura.com/ClassicMapper.asmx)
    ↓
DNS Resolution (hosts file or DNS server)
    ↓ (resolves to 127.0.0.1 or server IP)
    ↓
Flask Mock Backend (port 80)
    ↓
SOAP Request Parser
    ↓
CSV Data Files (./Data/)
    ↓
SOAP Response Builder
    ↓
    ← SOAP Response
PowerShell Client
```

## Security Note

⚠️ **This is a TESTING server only!**
- No authentication
- No encryption (HTTP, not HTTPS)
- No input validation
- CSV files are world-readable

**DO NOT use in production environments!**

## Support

For issues with the mock backend:
1. Check server console for error messages
2. Verify CSV files are properly formatted
3. Check DNS resolution with `nslookup`
4. Test with browser first before PowerShell

---

**Version:** 1.0  
**Last Updated:** 2025-10-13  
**Purpose:** Testing Desktop Management PowerShell implementation

