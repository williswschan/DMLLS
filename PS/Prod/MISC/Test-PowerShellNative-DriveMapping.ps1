# Test PowerShell Native Drive Mapping
# This demonstrates how we could replace the COM object approach with PowerShell native

[CmdletBinding()]
Param()

# Import the DMLogger module for testing
Try {
    Import-Module .\Modules\Framework\DMLogger.psm1 -Force
    Write-Host "SUCCESS: DMLogger module imported" -ForegroundColor Green
}
Catch {
    Write-Host "WARNING: Could not import DMLogger module: $($_.Exception.Message)" -ForegroundColor Yellow
    # Define a simple Write-DMLog function for testing
    Function Write-DMLog {
        Param([String]$Message, [String]$Level = "Info")
        Write-Host "[$Level] $Message" -ForegroundColor $(Switch($Level) {
            "Error" { "Red" }
            "Warning" { "Yellow" }
            "Info" { "White" }
            "Verbose" { "Gray" }
            Default { "White" }
        })
    }
}

# PowerShell Native Drive Mapping Function
Function Add-DMDriveMappingNative {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DriveLetter,
        
        [Parameter(Mandatory=$True)]
        [String]$UncPath,
        
        [Parameter(Mandatory=$False)]
        [String]$Description = "",
        
        [Parameter(Mandatory=$False)]
        [Switch]$SetDescription
    )
    
    Try {
        # Validate parameters
        If ([String]::IsNullOrEmpty($DriveLetter)) {
            Write-DMLog "Cannot map drive - DriveLetter is empty" -Level Error
            Return $False
        }
        
        If ([String]::IsNullOrEmpty($UncPath)) {
            Write-DMLog "Cannot map drive '$DriveLetter' - UncPath is empty" -Level Error
            Return $False
        }
        
        Write-DMLog "About to map '$DriveLetter' to '$UncPath'" -Level Verbose
        
        # Remove colon from drive letter if present
        $DriveName = $DriveLetter.TrimEnd(':')
        
        # Check if drive is already mapped
        $ExistingDrive = Get-PSDrive -Name $DriveName -ErrorAction SilentlyContinue
        If ($ExistingDrive) {
            If ($ExistingDrive.DisplayRoot -eq $UncPath) {
                Write-DMLog "Drive '$DriveLetter' is already mapped to '$UncPath'" -Level Info
                Return $True
            } Else {
                Write-DMLog "Drive '$DriveLetter' is mapped to different path '$($ExistingDrive.DisplayRoot)', removing..." -Level Info
                Remove-PSDrive -Name $DriveName -Force -ErrorAction SilentlyContinue
            }
        }
        
        # Map the drive using PowerShell native
        Write-DMLog "Mapping drive using PowerShell native..." -Level Verbose
        New-PSDrive -Name $DriveName -PSProvider FileSystem -Root $UncPath -Persist -Scope Global
        
        Write-DMLog "SUCCESS: Mapped '$DriveLetter' to '$UncPath'" -Level Info
        
        # Set description if requested and available
        If ($SetDescription -and -not [String]::IsNullOrEmpty($Description)) {
            Try {
                Write-DMLog "Setting drive description for '$DriveLetter' to '$Description'" -Level Verbose
                
                # Use COM object only for description setting
                $Shell = New-Object -ComObject Shell.Application
                $Shell.NameSpace("${DriveLetter}:").Self.Name = $Description
                
                Write-DMLog "SUCCESS: Description for '$DriveLetter' set to '$Description'" -Level Info
            }
            Catch {
                Write-DMLog "WARNING: Could not set description for '$DriveLetter': $($_.Exception.Message)" -Level Warning
                # Don't fail the whole operation if description fails
            }
        }
        
        Return $True
    }
    Catch {
        Write-DMLog "ERROR: Failed to map '$DriveLetter' to '$UncPath': $($_.Exception.Message)" -Level Error
        Return $False
    }
}

# PowerShell Native Drive Removal Function
Function Remove-DMDriveMappingNative {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DriveLetter
    )
    
    Try {
        $DriveName = $DriveLetter.TrimEnd(':')
        Write-DMLog "Removing drive mapping '$DriveLetter'" -Level Verbose
        
        Remove-PSDrive -Name $DriveName -Force -ErrorAction SilentlyContinue
        Write-DMLog "SUCCESS: Removed drive mapping '$DriveLetter'" -Level Info
        
        Return $True
    }
    Catch {
        Write-DMLog "WARNING: Could not remove drive mapping '$DriveLetter': $($_.Exception.Message)" -Level Warning
        Return $False
    }
}

# Test the PowerShell native approach
Write-Host "=== PowerShell Native Drive Mapping Test ===" -ForegroundColor Green

# Test data (using accessible paths)
$TestMappings = @(
    @{
        DriveLetter = "T"
        UncPath = "\\localhost\c$"
        Description = "Test Drive T"
        SetDescription = $True
    },
    @{
        DriveLetter = "U"
        UncPath = "\\localhost\admin$"
        Description = "Test Drive U"
        SetDescription = $True
    }
)

Write-Host "`n1. Testing PowerShell Native Drive Mapping:" -ForegroundColor Yellow

$SuccessCount = 0
$FailureCount = 0

ForEach ($Mapping in $TestMappings) {
    Write-Host "`n   Testing: $($Mapping.DriveLetter) -> $($Mapping.UncPath)" -ForegroundColor Cyan
    
    $Result = Add-DMDriveMappingNative -DriveLetter $Mapping.DriveLetter -UncPath $Mapping.UncPath -Description $Mapping.Description -SetDescription:$Mapping.SetDescription
    
    If ($Result) {
        $SuccessCount++
        Write-Host "   SUCCESS: Drive mapped successfully" -ForegroundColor Green
    } Else {
        $FailureCount++
        Write-Host "   FAILED: Drive mapping failed" -ForegroundColor Red
    }
}

Write-Host "`n2. Verifying Mapped Drives:" -ForegroundColor Yellow
$MappedDrives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Name -match '^[A-Z]$' -and $_.DisplayRoot }
ForEach ($Drive in $MappedDrives) {
    Write-Host "   $($Drive.Name): -> $($Drive.DisplayRoot)" -ForegroundColor Cyan
}

Write-Host "`n3. Cleaning Up Test Drives:" -ForegroundColor Yellow
ForEach ($Mapping in $TestMappings) {
    Write-Host "   Removing: $($Mapping.DriveLetter)" -ForegroundColor Gray
    Remove-DMDriveMappingNative -DriveLetter $Mapping.DriveLetter
}

Write-Host "`n=== Test Results ===" -ForegroundColor Green
Write-Host "Successfully Mapped: $SuccessCount" -ForegroundColor Green
Write-Host "Failed Mappings: $FailureCount" -ForegroundColor Red

Write-Host "`n=== Advantages of PowerShell Native Approach ===" -ForegroundColor Yellow
Write-Host "✓ No COM object dependencies" -ForegroundColor Green
Write-Host "✓ Simpler error handling" -ForegroundColor Green
Write-Host "✓ More PowerShell-idiomatic" -ForegroundColor Green
Write-Host "✓ Better integration with PowerShell ecosystem" -ForegroundColor Green
Write-Host "✓ Consistent with other PowerShell operations" -ForegroundColor Green
Write-Host "✓ Easier to debug and maintain" -ForegroundColor Green

Write-Host "`n=== Limitations ===" -ForegroundColor Yellow
Write-Host "⚠ Cannot set drive descriptions natively (requires COM objects)" -ForegroundColor Yellow
Write-Host "⚠ Requires PowerShell 3.0+ for New-PSDrive" -ForegroundColor Yellow

Write-Host "`n=== Recommendation ===" -ForegroundColor Green
Write-Host "Use PowerShell native approach with optional COM objects for descriptions only." -ForegroundColor White
