<#
.SYNOPSIS
    Desktop Management User Session Inventory Module
    
.DESCRIPTION
    Tracks user logon and logoff events by sending session data to the backend.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: InventoryUserSessionLogon_W10.vbs, InventoryUserSessionLogoff_W10.vbs
#>

# Import required modules
Using Module ..\Services\DMInventoryService.psm1
Using Module ..\Framework\DMLogger.psm1

<#
.SYNOPSIS
    Sends user session logon inventory to backend.
    
.DESCRIPTION
    Records user logon event with computer and user information.
    This runs during user logon to track who logged in, when, and where.
    
.PARAMETER UserInfo
    User information object from Get-DMUserInfo
    
.PARAMETER ComputerInfo
    Computer information object from Get-DMComputerInfo
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Invoke-DMUserSessionLogonInventory -UserInfo $User -ComputerInfo $Computer
#>
Function Invoke-DMUserSessionLogonInventory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo
    )
    
    Try {
        Write-DMLog "User Session Logon Inventory: Starting" -Level Verbose
        
        # Log what we're about to send
        Write-DMLog "User Session Logon Inventory: About to insert following data:" -Level Verbose
        Write-DMLog "  User Name: $($UserInfo.Name)" -Level Verbose
        Write-DMLog "  User Domain: $($UserInfo.Domain)" -Level Verbose
        Write-DMLog "  User OuMapping: $($UserInfo.OUMapping)" -Level Verbose
        Write-DMLog "  Computer Name: $($ComputerInfo.Name)" -Level Verbose
        Write-DMLog "  Computer Domain: $($ComputerInfo.Domain)" -Level Verbose
        Write-DMLog "  Computer Site: $($ComputerInfo.Site)" -Level Verbose
        Write-DMLog "  Computer CityCode: $($ComputerInfo.CityCode)" -Level Verbose
        
        # Send logon inventory
        [Boolean]$Success = Send-DMLogonInventory -UserInfo $UserInfo -ComputerInfo $ComputerInfo
        
        If ($Success) {
            Write-DMLog "User Session Logon Inventory: Successfully sent session data" -Level Verbose
        } Else {
            Write-DMLog "User Session Logon Inventory: Failed to send session data" -Level Warning
        }
        
        Write-DMLog "User Session Logon Inventory: Completed" -Level Verbose
        Return $Success
    }
    Catch {
        Write-DMLog "User Session Logon Inventory: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Sends user session logoff inventory to backend.
    
.DESCRIPTION
    Records user logoff event.
    This runs during user logoff to track when the user logged off.
    
.PARAMETER UserInfo
    User information object from Get-DMUserInfo
    
.PARAMETER ComputerInfo
    Computer information object from Get-DMComputerInfo
    
.OUTPUTS
    Boolean - true if successful
    
.EXAMPLE
    $Success = Invoke-DMUserSessionLogoffInventory -UserInfo $User -ComputerInfo $Computer
#>
Function Invoke-DMUserSessionLogoffInventory {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo
    )
    
    Try {
        Write-DMLog "User Session Logoff Inventory: Starting" -Level Verbose
        
        # Log what we're about to send
        Write-DMLog "User Session Logoff Inventory: About to insert following data:" -Level Verbose
        Write-DMLog "  User Name: $($UserInfo.Name)" -Level Verbose
        Write-DMLog "  User Domain: $($UserInfo.Domain)" -Level Verbose
        Write-DMLog "  User OuMapping: $($UserInfo.OUMapping)" -Level Verbose
        Write-DMLog "  Computer Name: $($ComputerInfo.Name)" -Level Verbose
        Write-DMLog "  Computer Domain: $($ComputerInfo.Domain)" -Level Verbose
        Write-DMLog "  Computer Site: $($ComputerInfo.Site)" -Level Verbose
        Write-DMLog "  Computer CityCode: $($ComputerInfo.CityCode)" -Level Verbose
        
        # Send logoff inventory
        [Boolean]$Success = Send-DMLogoffInventory -UserInfo $UserInfo -ComputerInfo $ComputerInfo
        
        If ($Success) {
            Write-DMLog "User Session Logoff Inventory: Successfully sent session data" -Level Verbose
        } Else {
            Write-DMLog "User Session Logoff Inventory: Failed to send session data" -Level Warning
        }
        
        Write-DMLog "User Session Logoff Inventory: Completed" -Level Verbose
        Return $Success
    }
    Catch {
        Write-DMLog "User Session Logoff Inventory: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Invoke-DMUserSessionLogonInventory',
    'Invoke-DMUserSessionLogoffInventory'
)

