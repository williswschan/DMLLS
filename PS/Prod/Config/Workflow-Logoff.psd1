@{
    # Desktop Management Suite - Logoff Workflow Configuration
    # Defines the sequence of actions to execute during Windows 10 user logoff
    
    JobType = 'Logoff'
    Description = 'Windows 10 Desktop User Logoff'
    
    # Workflow steps (executed in order)
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
            Description = 'Track user logoff event in backend database'
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
        },
        
        @{
            Order = 120
            Name = 'Printer Inventory'
            Phase = 'Inventory'
            Module = 'Inventory\Invoke-PrinterInventory.psm1'
            Function = 'Invoke-DMPrinterInventory'
            Enabled = $True
            Parameters = @{
                UserInfo = 'UserInfo'
                ComputerInfo = 'ComputerInfo'
            }
            ContinueOnError = $True
            Description = 'Send current printer mappings to backend'
        },
        
        @{
            Order = 130
            Name = 'PST Inventory'
            Phase = 'Inventory'
            Module = 'Inventory\Invoke-PersonalFolderInventory.psm1'
            Function = 'Invoke-DMPersonalFolderInventory'
            Enabled = $True
            Parameters = @{
                UserInfo = 'UserInfo'
                ComputerInfo = 'ComputerInfo'
            }
            ContinueOnError = $True
            Description = 'Send current PST file locations to backend'
        },
        
        # ====================================================================
        # PHASE 2: UTILITIES
        # ====================================================================
        
        @{
            Order = 200
            Name = 'Power Configuration Revert'
            Phase = 'Utilities'
            Module = 'Utilities\Set-PowerConfiguration.psm1'
            Function = 'Set-DMPowerConfiguration'
            Enabled = $True
            Parameters = @{
                JobType = 'Static:Logoff'  # Static value
            }
            ContinueOnError = $True
            Description = 'Revert monitor timeout to default (20 minutes)'
        }
    )
}

