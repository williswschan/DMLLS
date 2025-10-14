@{
    # Desktop Management Suite - Logon Workflow Configuration
    # Defines the sequence of actions to execute during Windows 10 user logon
    
    JobType = 'Logon'
    Description = 'Windows 10 Desktop User Logon'
    
    # Workflow steps (executed in order)
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
            Description = 'Track user logon event in backend database'
        },
        
        # ====================================================================
        # PHASE 2: MAPPER (Get FROM Backend and Apply)
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
            Description = 'Map network drives based on backend configuration'
            # Note: Invoke-DriveMapper depends on functions from DriveInventory
        },
        
        @{
            Order = 210
            Name = 'Printer Mapper'
            Phase = 'Mapper'
            Module = 'Mapper\Invoke-PrinterMapper.psm1'
            Function = 'Invoke-DMPrinterMapper'
            Enabled = $True
            Parameters = @{
                UserInfo = 'UserInfo'
                ComputerInfo = 'ComputerInfo'
            }
            ContinueOnError = $True
            Description = 'Map network printers based on backend configuration'
        },
        
        @{
            Order = 220
            Name = 'PST Mapper'
            Phase = 'Mapper'
            Module = 'Mapper\Invoke-PersonalFolderMapper.psm1'
            Function = 'Invoke-DMPersonalFolderMapper'
            Enabled = $True
            Parameters = @{
                UserInfo = 'UserInfo'
                ComputerInfo = 'ComputerInfo'
            }
            ContinueOnError = $True
            Description = 'Map Outlook PST files based on backend configuration'
        },
        
        # ====================================================================
        # PHASE 3: UTILITIES
        # ====================================================================
        
        @{
            Order = 300
            Name = 'Power Configuration'
            Phase = 'Utilities'
            Module = 'Utilities\Set-PowerConfiguration.psm1'
            Function = 'Set-DMPowerConfiguration'
            Enabled = $True
            Parameters = @{
                JobType = 'Static:Logon'  # Static value, not from context
            }
            ContinueOnError = $True
            Description = 'Configure monitor timeout based on screensaver GPO'
        },
        
        @{
            Order = 310
            Name = 'Password Expiry Notification'
            Phase = 'Utilities'
            Module = 'Utilities\Show-PasswordExpiryNotification.psm1'
            Function = 'Show-DMPasswordExpiryNotification'
            Enabled = $True
            Parameters = @{
                UserInfo = 'UserInfo'
            }
            ContinueOnError = $True
            Description = 'Show password expiry warning if within 14 days'
        },
        
        @{
            Order = 320
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
            Description = 'Import IE zone settings from registry file (LEGACY)'
        },
        
        @{
            Order = 330
            Name = 'Retail Home Drive Label'
            Phase = 'Utilities'
            Module = 'Utilities\Set-RetailHomeDriveLabel.psm1'
            Function = 'Set-DMRetailHomeDriveLabel'
            Enabled = $True
            Parameters = @{
                UserInfo = 'UserInfo'
            }
            ContinueOnError = $True
            Description = 'Set V: drive label for Retail users'
        }
    )
}

