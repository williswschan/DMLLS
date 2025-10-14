@{
    # Desktop Management Suite - Terminal Server Logon Workflow Configuration
    # Defines the sequence of actions for Terminal Server/Citrix user logon
    
    JobType = 'TSLogon'
    Description = 'Terminal Server/Citrix User Logon'
    
    # Workflow steps (executed in order)
    # Note: Terminal Server has simplified workflow (drives only, no printers/PST)
    Steps = @(
        # ====================================================================
        # PHASE 1: INVENTORY (Send TO Backend)
        # ====================================================================
        
        @{
            Order = 100
            Name = 'User Session Logon Inventory'
            Phase = 'Inventory'
            Module = 'Inventory\Invoke-UserSessionInventory.psm1'
            Function = 'Invoke-DMUserSessionLogonInventory'
            Enabled = $True
            Parameters = @{
                UserInfo = 'UserInfo'
                ComputerInfo = 'ComputerInfo'
            }
            ContinueOnError = $True
            Description = 'Track Terminal Server user logon event'
        },
        
        # ====================================================================
        # PHASE 2: MAPPER (Drives Only for Terminal Server)
        # ====================================================================
        
        @{
            Order = 200
            Name = 'Drive Mapper'
            Phase = 'Mapper'
            Module = 'Mapper\Invoke-DriveMapper.psm1'
            Function = 'Invoke-DMDriveMapper'
            Enabled = $True
            Parameters = @{
                UserInfo = 'UserInfo'
                ComputerInfo = 'ComputerInfo'
            }
            ContinueOnError = $True
            Description = 'Map network drives (Terminal Server - drives only)'
            # Note: Invoke-DriveMapper depends on functions from DriveInventory
        },
        
        # ====================================================================
        # PHASE 3: UTILITIES (Minimal for Terminal Server)
        # ====================================================================
        
        @{
            Order = 300
            Name = 'IE Zone Configuration'
            Phase = 'Utilities'
            Module = 'Utilities\Import-IEZoneConfiguration.psm1'
            Function = 'Import-DMIEZoneConfiguration'
            Enabled = $False  # Disabled by default (IE is deprecated)
            Parameters = @{
                JobType = 'Static:Logon'
                UserInfo = 'UserInfo'
            }
            ContinueOnError = $True
            Description = 'Import IE zone settings (LEGACY)'
        }
    )
}

