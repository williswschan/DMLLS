<#
.SYNOPSIS
    Desktop Management Retail Home Drive Label Module
    
.DESCRIPTION
    Changes the V: drive label for Retail users.
    Replaces default "documents" label with Japanese label containing username.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: SetRetailHomeDriveLabel.vbs
#>

# Import required modules
Using Module .\Test-Environment.psm1
Using Module ..\Framework\DMLogger.psm1

<#
.SYNOPSIS
    Sets Retail home drive (V:) label.
    
.DESCRIPTION
    Changes V: drive label from default "documents" to "'Fs' 名 <USERNAME>".
    Only runs for Retail users who have V: drive.
    
.PARAMETER UserInfo
    User information object
    
.OUTPUTS
    Boolean - true if successful or not applicable
    
.EXAMPLE
    Set-DMRetailHomeDriveLabel -UserInfo $User
#>
Function Set-DMRetailHomeDriveLabel {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo
    )
    
    Try {
        Write-DMLog "Retail HomeDrive: Starting to change label on homedrive" -Level Info
        
        # Check if user is Retail
        [Boolean]$IsRetail = Test-DMRetailUser -DistinguishedName $UserInfo.DistinguishedName
        
        If (-not $IsRetail) {
            Write-DMLog "Retail HomeDrive: The user does not belong to Retail OU. Terminating script execution" -Level Info
            Write-DMLog "Retail HomeDrive: Completed" -Level Info
            Return $True
        }
        
        # Check if V: drive exists
        If (-not (Test-Path -Path "V:\")) {
            Write-DMLog "Retail HomeDrive: V drive does not exist. Terminating script execution" -Level Info
            Write-DMLog "Retail HomeDrive: Completed" -Level Info
            Return $True
        }
        
        # Get current label using Shell.Application
        [Object]$Shell = New-Object -ComObject Shell.Application
        [Object]$VDrive = $Shell.NameSpace("V:")
        
        If ($Null -eq $VDrive) {
            Write-DMLog "Retail HomeDrive: Cannot access V: drive namespace" -Level Warning
            Return $False
        }
        
        [String]$CurrentLabel = $VDrive.Self.Name
        Write-DMLog "Retail HomeDrive: Current label is '$CurrentLabel'" -Level Info
        
        # Check if current label starts with "documents" (default label)
        If ($CurrentLabel.ToLower().StartsWith("documents")) {
            # Change to new label: 'Fs' 名 <USERNAME>
            # Note: 名 is Japanese character for "name"
            [String]$NewLabel = "'Fs' 名 $env:USERNAME"
            
            Try {
                $VDrive.Self.Name = $NewLabel
                Write-DMLog "Retail HomeDrive: New label is now '$NewLabel'" -Level Info
            } Catch {
                Write-DMLog "Retail HomeDrive: Failed to set new label: $($_.Exception.Message)" -Level Warning
                Return $False
            }
        } Else {
            Write-DMLog "Retail HomeDrive: Change is not required" -Level Info
        }
        
        Write-DMLog "Retail HomeDrive: Completed" -Level Info
        Return $True
    }
    Catch {
        Write-DMLog "Retail HomeDrive: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Set-DMRetailHomeDriveLabel'
)

