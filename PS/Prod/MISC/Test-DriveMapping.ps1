# Test Drive Mapping - Simple demonstration
# This script shows exactly how drive mapping works

[CmdletBinding()]
Param()

Write-Host "=== Drive Mapping Test ===" -ForegroundColor Green

# Test 1: Show current mapped drives
Write-Host "`n1. Current Mapped Drives:" -ForegroundColor Yellow
Try {
    $Network = New-Object -ComObject WScript.Network
    $MappedDrives = $Network.EnumNetworkDrives()
    
    For ($i = 0; $i -lt $MappedDrives.Count; $i += 2) {
        $DriveLetter = $MappedDrives.Item($i)
        $UncPath = $MappedDrives.Item($i + 1)
        Write-Host "   $DriveLetter -> $UncPath" -ForegroundColor Cyan
    }
}
Catch {
    Write-Host "   Error enumerating drives: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 2: Try to map a test drive
Write-Host "`n2. Testing Drive Mapping:" -ForegroundColor Yellow

# Use a simple test path (adjust this to a path you have access to)
$TestDrive = "T:"
$TestPath = "\\localhost\c$"  # This should work on most Windows systems

Write-Host "   Attempting to map $TestDrive to $TestPath" -ForegroundColor Cyan

Try {
    $Network = New-Object -ComObject WScript.Network
    
    # Check if drive is already mapped
    $MappedDrives = $Network.EnumNetworkDrives()
    $AlreadyMapped = $False
    
    For ($i = 0; $i -lt $MappedDrives.Count; $i += 2) {
        If ($MappedDrives.Item($i) -eq $TestDrive) {
            $AlreadyMapped = $True
            Write-Host "   Drive $TestDrive is already mapped to $($MappedDrives.Item($i + 1))" -ForegroundColor Yellow
            Break
        }
    }
    
    If (-not $AlreadyMapped) {
        Write-Host "   Calling MapNetworkDrive..." -ForegroundColor Cyan
        $Network.MapNetworkDrive($TestDrive, $TestPath, $True)
        Write-Host "   SUCCESS: Drive mapped successfully!" -ForegroundColor Green
        
        # Try to set description
        Try {
            Write-Host "   Setting drive description..." -ForegroundColor Cyan
            $Shell = New-Object -ComObject Shell.Application
            $Shell.NameSpace($TestDrive).Self.Name = "Test Drive"
            Write-Host "   SUCCESS: Description set!" -ForegroundColor Green
        }
        Catch {
            Write-Host "   WARNING: Could not set description: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}
Catch {
    Write-Host "   ERROR: Failed to map drive: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "   Error Details: $($_.Exception)" -ForegroundColor Red
}

# Test 3: Show mapped drives again
Write-Host "`n3. Mapped Drives After Test:" -ForegroundColor Yellow
Try {
    $MappedDrives = $Network.EnumNetworkDrives()
    
    For ($i = 0; $i -lt $MappedDrives.Count; $i += 2) {
        $DriveLetter = $MappedDrives.Item($i)
        $UncPath = $MappedDrives.Item($i + 1)
        Write-Host "   $DriveLetter -> $UncPath" -ForegroundColor Cyan
    }
}
Catch {
    Write-Host "   Error enumerating drives: $($_.Exception.Message)" -ForegroundColor Red
}

# Test 4: Clean up test drive
Write-Host "`n4. Cleaning up test drive:" -ForegroundColor Yellow
Try {
    $Network.RemoveNetworkDrive($TestDrive, $True, $True)
    Write-Host "   SUCCESS: Test drive removed" -ForegroundColor Green
}
Catch {
    Write-Host "   Could not remove test drive: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`n=== Test Complete ===" -ForegroundColor Green
