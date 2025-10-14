@{
    # Desktop Management Suite - Terminal Server Logoff Workflow Configuration
    # Defines the sequence of actions for Terminal Server/Citrix user logoff
    
    JobType = 'TSLogoff'
    Description = 'Terminal Server/Citrix User Logoff'
    
    # Workflow steps (executed in order)
    # Note: Minimal workflow for Terminal Server logoff
    Steps = @(
        # ====================================================================
        # PHASE 1: INVENTORY (Send TO Backend)
        # ====================================================================
        
        @{
            Order = 100
            Name = 'User Session Logoff Inventory'
            Phase = 'Inventory'
            Module = 'Inventory\Invoke-UserSessionInventory.psm1'
            Function = 'Invoke-DMUserSessionLogoffInventory'
            Enabled = $True
            Parameters = @{
                UserInfo = 'UserInfo'
                ComputerInfo = 'ComputerInfo'
            }
            ContinueOnError = $True
            Description = 'Track Terminal Server user logoff event'
        },
        
        @{
            Order = 110
            Name = 'Drive Inventory'
            Phase = 'Inventory'
            Module = 'Inventory\Invoke-DriveInventory.psm1'
            Function = 'Invoke-DMDriveInventory'
            Enabled = $True
            Parameters = @{
                UserInfo = 'UserInfo'
                ComputerInfo = 'ComputerInfo'
            }
            ContinueOnError = $True
            Description = 'Send current drive mappings to backend'
        }
    )
}

