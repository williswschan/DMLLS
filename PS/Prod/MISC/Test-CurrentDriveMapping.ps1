# Test Current Drive Mapping Implementation
# This script tests our current drive mapping logic step by step

[CmdletBinding()]
Param()

Write-Host "=== Current Drive Mapping Implementation Test ===" -ForegroundColor Green

# Simulate the drive mapping data that would come from the service
$TestMappings = @(
    [PSCustomObject]@{
        PSTypeName = 'DM.DriveMapping'
        Id = "1"
        Domain = "TESTDOMAIN"
        UserId = "testuser"
        AdGroup = "TestGroup"
        Site = "TESTSITE"
        DriveLetter = "H"
        UncPath = "\\rndasiapac.nom\home\userdata\HKG\chanwilw\Documents"
        Description = "Home Drive"
        DisconnectOnLogin = $False
    },
    [PSCustomObject]@{
        PSTypeName = 'DM.DriveMapping'
        Id = "2"
        Domain = "TESTDOMAIN"
        UserId = "testuser"
        AdGroup = "TestGroup"
        Site = "TESTSITE"
        DriveLetter = "Y"
        UncPath = "\\10.190.131.70\c$"
        Description = "Test Drive Y"
        DisconnectOnLogin = $False
    },
    [PSCustomObject]@{
        PSTypeName = 'DM.DriveMapping'
        Id = "3"
        Domain = "TESTDOMAIN"
        UserId = "testuser"
        AdGroup = "TestGroup"
        Site = "TESTSITE"
        DriveLetter = "Z"
        UncPath = "\\10.190.131.14\c$"
        Description = "Test Drive Z"
        DisconnectOnLogin = $False
    }
)

Write-Host "`n1. Testing Drive Mapping Data:" -ForegroundColor Yellow
ForEach ($Mapping in $TestMappings) {
    Write-Host "   Drive: $($Mapping.DriveLetter) -> $($Mapping.UncPath)" -ForegroundColor Cyan
    Write-Host "      Description: $($Mapping.Description)" -ForegroundColor Gray
    Write-Host "      DisconnectOnLogin: $($Mapping.DisconnectOnLogin)" -ForegroundColor Gray
}

Write-Host "`n2. Testing Parameter Processing:" -ForegroundColor Yellow
ForEach ($Mapping in $TestMappings) {
    Write-Host "`n   Processing Drive: $($Mapping.DriveLetter)" -ForegroundColor Cyan
    
    # Test the same logic as in Set-DMDriveMapping
    Try {
        # Validate mapping has required properties
        If ($Null -eq $Mapping -or [String]::IsNullOrEmpty($Mapping.DriveLetter)) {
            Write-Host "      SKIP: Empty or null DriveLetter" -ForegroundColor Red
            Continue
        }
        
        [String]$DriveLetter = $Mapping.DriveLetter.TrimEnd(':')
        [String]$UncPath = If (-not [String]::IsNullOrEmpty($Mapping.UncPath)) { 
            [System.Environment]::ExpandEnvironmentVariables($Mapping.UncPath) 
        } Else { "" }
        [String]$Description = If (-not [String]::IsNullOrEmpty($Mapping.Description)) { $Mapping.Description } Else { "" }
        
        Write-Host "      Processed DriveLetter: '$DriveLetter'" -ForegroundColor Green
        Write-Host "      Processed UncPath: '$UncPath'" -ForegroundColor Green
        Write-Host "      Processed Description: '$Description'" -ForegroundColor Green
        
        # Test COM object creation
        Write-Host "      Testing COM object creation..." -ForegroundColor Gray
        Try {
            $Network = New-Object -ComObject WScript.Network
            Write-Host "      SUCCESS: WScript.Network COM object created" -ForegroundColor Green
        }
        Catch {
            Write-Host "      ERROR: Failed to create WScript.Network COM object: $($_.Exception.Message)" -ForegroundColor Red
            Continue
        }
        
        # Test Shell COM object creation
        Try {
            $Shell = New-Object -ComObject Shell.Application
            Write-Host "      SUCCESS: Shell.Application COM object created" -ForegroundColor Green
        }
        Catch {
            Write-Host "      WARNING: Failed to create Shell.Application COM object: $($_.Exception.Message)" -ForegroundColor Yellow
        }
        
        # Test actual mapping (but don't actually map - just test the call)
        Write-Host "      Testing MapNetworkDrive call (dry run)..." -ForegroundColor Gray
        Try {
            # We won't actually call MapNetworkDrive to avoid mapping test drives
            # But we can test if the parameters are valid
            If ([String]::IsNullOrEmpty($DriveLetter)) {
                Write-Host "      ERROR: DriveLetter is empty" -ForegroundColor Red
            }
            ElseIf ([String]::IsNullOrEmpty($UncPath)) {
                Write-Host "      ERROR: UncPath is empty" -ForegroundColor Red
            }
            Else {
                Write-Host "      SUCCESS: Parameters look valid for mapping" -ForegroundColor Green
                Write-Host "         Would call: Network.MapNetworkDrive('$DriveLetter:', '$UncPath', True)" -ForegroundColor Gray
            }
        }
        Catch {
            Write-Host "      ERROR: Failed to validate mapping parameters: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    Catch {
        Write-Host "      ERROR: Failed to process mapping: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "      Error Details: $($_.Exception)" -ForegroundColor Red
    }
}

Write-Host "`n3. Testing Current Mapped Drives:" -ForegroundColor Yellow
Try {
    $Network = New-Object -ComObject WScript.Network
    $MappedDrives = $Network.EnumNetworkDrives()
    
    If ($MappedDrives.Count -eq 0) {
        Write-Host "   No drives currently mapped" -ForegroundColor Gray
    }
    Else {
        For ($i = 0; $i -lt $MappedDrives.Count; $i += 2) {
            $DriveLetter = $MappedDrives.Item($i)
            $UncPath = $MappedDrives.Item($i + 1)
            Write-Host "   $DriveLetter -> $UncPath" -ForegroundColor Cyan
        }
    }
}
Catch {
    Write-Host "   Error enumerating drives: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n=== Test Complete ===" -ForegroundColor Green
Write-Host "`nTo test actual drive mapping, run: .\Test-DriveMapping.ps1" -ForegroundColor Yellow
