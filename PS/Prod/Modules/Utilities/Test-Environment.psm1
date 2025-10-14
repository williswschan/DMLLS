<#
.SYNOPSIS
    Desktop Management Environment Detection Module
    
.DESCRIPTION
    Detects VPN connection, Retail users/computers, Terminal sessions, VDI environments.
    Replacement for detection functions scattered across VBScript modules.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: RetailCommon.vbs, VPN detection, session detection functions
#>

<#
.SYNOPSIS
    Tests if Cisco AnyConnect VPN is connected.
    
.DESCRIPTION
    Checks if Cisco AnyConnect VPN adapter is present and connected.
    Uses Get-NetAdapter instead of WMI for better performance.
    
.OUTPUTS
    Boolean - true if VPN is connected
    
.EXAMPLE
    If (Test-DMVPNConnection) { ... }
#>
Function Test-DMVPNConnection {
    [CmdletBinding()]
    Param()
    
    Try {
        # Check for Cisco AnyConnect network adapter that is Up
        [Array]$VPNAdapters = Get-NetAdapter -ErrorAction SilentlyContinue | 
            Where-Object { 
                $_.Name -like '*Cisco*' -and 
                $_.Status -eq 'Up' 
            }
        
        If ($VPNAdapters.Count -gt 0) {
            Return $True
        }
        
        # Fallback: Check using WMI (for compatibility)
        [Array]$VPNWMIAdapters = Get-CimInstance -ClassName Win32_NetworkAdapter -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -like '*Cisco AnyConnect*' -and
                $_.NetConnectionStatus -eq 2  # 2 = Connected
            }
        
        Return ($VPNWMIAdapters.Count -gt 0)
    }
    Catch {
        Return $False
    }
}

<#
.SYNOPSIS
    Tests if a user is a Retail user.
    
.DESCRIPTION
    Checks if user DN contains Retail organizational units.
    Retail OUs: "OU=Nomura Retail", "OU=Nomura Trust Bank", "OU=TOK,OU=Nomura Asset Management"
    
.PARAMETER DistinguishedName
    User's LDAP Distinguished Name
    
.OUTPUTS
    Boolean - true if Retail user
    
.EXAMPLE
    $IsRetail = Test-DMRetailUser -DistinguishedName $User.DistinguishedName
#>
Function Test-DMRetailUser {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [AllowEmptyString()]
        [String]$DistinguishedName = ""
    )
    
    If ([String]::IsNullOrEmpty($DistinguishedName)) {
        Return $False
    }
    
    # Check for Retail OU patterns
    [Array]$RetailPatterns = @(
        'OU=Nomura Retail',
        'OU=Nomura Trust Bank',
        'OU=TOK,OU=Nomura Asset Management'
    )
    
    ForEach ($Pattern in $RetailPatterns) {
        If ($DistinguishedName -like "*$Pattern*") {
            Return $True
        }
    }
    
    Return $False
}

<#
.SYNOPSIS
    Tests if a computer is a Retail computer.
    
.DESCRIPTION
    Checks if computer DN contains Retail organizational units.
    Same patterns as Retail user detection.
    
.PARAMETER DistinguishedName
    Computer's LDAP Distinguished Name
    
.OUTPUTS
    Boolean - true if Retail computer
    
.EXAMPLE
    $IsRetail = Test-DMRetailComputer -DistinguishedName $Computer.DistinguishedName
#>
Function Test-DMRetailComputer {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [AllowEmptyString()]
        [String]$DistinguishedName = ""
    )
    
    # Same logic as user detection
    Return Test-DMRetailUser -DistinguishedName $DistinguishedName
}

<#
.SYNOPSIS
    Tests if user is in a specific Retail group using whoami.
    
.DESCRIPTION
    Uses whoami /groups command to check group membership.
    Faster than LDAP queries for current user.
    
.PARAMETER GroupName
    Group name to check
    
.OUTPUTS
    Boolean - true if user is in group
    
.EXAMPLE
    $IsMember = Test-DMRetailUserGroup -GroupName "Retail Users"
#>
Function Test-DMRetailUserGroup {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$GroupName
    )
    
    Try {
        # Execute whoami /groups
        [Object]$Result = & whoami.exe /groups 2>&1
        
        If ($LASTEXITCODE -eq 0) {
            [String]$Output = $Result -join "`n"
            Return ($Output -like "*$GroupName*")
        }
        
        Return $False
    }
    Catch {
        Return $False
    }
}

<#
.SYNOPSIS
    Tests if running in a Terminal Server session.
    
.DESCRIPTION
    Checks SESSIONNAME environment variable and session type.
    Console sessions vs RDP/Citrix sessions.
    
.OUTPUTS
    PSCustomObject with IsTerminalSession, SessionType, SessionName properties
    
.EXAMPLE
    $Session = Test-DMTerminalSession
    If ($Session.IsTerminalSession) { ... }
#>
Function Test-DMTerminalSession {
    [CmdletBinding()]
    Param()
    
    [String]$SessionName = $env:SESSIONNAME
    [Boolean]$IsTerminal = $False
    [String]$SessionType = "Unknown"
    
    If ([String]::IsNullOrEmpty($SessionName)) {
        $SessionType = "Unknown"
    }
    ElseIf ($SessionName -like "Console*") {
        $SessionType = "Console"
        $IsTerminal = $False
    }
    ElseIf ($SessionName -like "ICA*") {
        $SessionType = "Citrix"
        $IsTerminal = $True
    }
    ElseIf ($SessionName -like "RDP*") {
        $SessionType = "RDP"
        $IsTerminal = $True
    }
    Else {
        # Assume terminal if not console
        $SessionType = "Terminal"
        $IsTerminal = $True
    }
    
    Return [PSCustomObject]@{
        PSTypeName = 'DM.SessionInfo'
        IsTerminalSession = $IsTerminal
        SessionType = $SessionType
        SessionName = $SessionName
    }
}

<#
.SYNOPSIS
    Tests if running on a Shared VDI (Retail-specific).
    
.DESCRIPTION
    Checks if hostname contains "JPRWV1" or "JPRWV3" patterns.
    Used for Retail VDI detection.
    
.PARAMETER ComputerName
    Computer name to check (defaults to current computer)
    
.OUTPUTS
    Boolean - true if Shared VDI
    
.EXAMPLE
    $IsSharedVDI = Test-DMSharedVDI
#>
Function Test-DMSharedVDI {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$ComputerName = $env:COMPUTERNAME
    )
    
    [Array]$SharedVDIPatterns = @('JPRWV1', 'JPRWV3')
    
    ForEach ($Pattern in $SharedVDIPatterns) {
        If ($ComputerName -like "*$Pattern*") {
            Return $True
        }
    }
    
    Return $False
}

<#
.SYNOPSIS
    Tests if running on a Virtual Machine.
    
.DESCRIPTION
    Checks computer model and manufacturer for VM indicators.
    Detects VMware, Hyper-V, VirtualBox, etc.
    
.OUTPUTS
    PSCustomObject with IsVirtual, Platform properties
    
.EXAMPLE
    $VM = Test-DMVirtualMachine
    If ($VM.IsVirtual) { ... }
#>
Function Test-DMVirtualMachine {
    [CmdletBinding()]
    Param()
    
    Try {
        [Object]$ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        
        [String]$Model = $ComputerSystem.Model
        [String]$Manufacturer = $ComputerSystem.Manufacturer
        
        [Boolean]$IsVirtual = $False
        [String]$Platform = "Physical"
        
        # Check for VM indicators
        If ($Model -like "*Virtual Machine*") {
            $IsVirtual = $True
            $Platform = "Hyper-V"
        }
        ElseIf ($Model -like "*VMware*") {
            $IsVirtual = $True
            $Platform = "VMware"
        }
        ElseIf ($Manufacturer -like "*VMware*") {
            $IsVirtual = $True
            $Platform = "VMware"
        }
        ElseIf ($Manufacturer -like "*Microsoft*" -and $Model -like "*Virtual*") {
            $IsVirtual = $True
            $Platform = "Hyper-V"
        }
        ElseIf ($Model -like "*VirtualBox*") {
            $IsVirtual = $True
            $Platform = "VirtualBox"
        }
        
        Return [PSCustomObject]@{
            PSTypeName = 'DM.VirtualMachineInfo'
            IsVirtual = $IsVirtual
            Platform = $Platform
            Model = $Model
            Manufacturer = $Manufacturer
        }
    }
    Catch {
        Return [PSCustomObject]@{
            PSTypeName = 'DM.VirtualMachineInfo'
            IsVirtual = $False
            Platform = "Unknown"
            Model = ""
            Manufacturer = ""
        }
    }
}

<#
.SYNOPSIS
    Tests if computer is a server.
    
.DESCRIPTION
    Checks OS caption and domain role to determine if server OS.
    
.OUTPUTS
    Boolean - true if server
    
.EXAMPLE
    $IsServer = Test-DMServerOS
#>
Function Test-DMServerOS {
    [CmdletBinding()]
    Param()
    
    Try {
        [Object]$OS = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        [Object]$Computer = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
        
        # Check OS caption for "Server"
        If ($OS.Caption -like "*Server*") {
            Return $True
        }
        
        # Check domain role (4-5 = Server)
        # 0 = Standalone Workstation, 1 = Member Workstation, 2 = Standalone Server
        # 3 = Member Server, 4 = Backup DC, 5 = Primary DC
        If ($Computer.DomainRole -ge 3) {
            Return $True
        }
        
        Return $False
    }
    Catch {
        Return $False
    }
}

<#
.SYNOPSIS
    Gets password change hotkey based on session type.
    
.DESCRIPTION
    Returns the appropriate hotkey for changing password based on session type.
    Used for password expiry notifications.
    
.PARAMETER SessionInfo
    Session info from Test-DMTerminalSession
    
.OUTPUTS
    String - hotkey combination
    
.EXAMPLE
    $Session = Test-DMTerminalSession
    $Hotkey = Get-DMPasswordChangeHotkey -SessionInfo $Session
#>
Function Get-DMPasswordChangeHotkey {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [PSCustomObject]$SessionInfo
    )
    
    If ($Null -eq $SessionInfo) {
        $SessionInfo = Test-DMTerminalSession
    }
    
    # Determine hotkey based on session type
    Switch ($SessionInfo.SessionType) {
        'Citrix' { Return 'CTRL + F1' }
        'RDP'    { Return 'CTRL + ALT + END' }
        'Console' {
            # Check if VDI
            [Object]$VMInfo = Test-DMVirtualMachine
            If ($VMInfo.IsVirtual) {
                Return 'CTRL + ALT + INS'
            } Else {
                Return 'CTRL + ALT + DEL'
            }
        }
        Default  { Return 'CTRL + ALT + DEL' }
    }
}

<#
.SYNOPSIS
    Gets system language/locale.
    
.DESCRIPTION
    Returns the current system locale (e.g., "ja-JP", "en-US").
    Used for multi-language password notifications.
    
.OUTPUTS
    String - locale code
    
.EXAMPLE
    $Locale = Get-DMSystemLocale
#>
Function Get-DMSystemLocale {
    [CmdletBinding()]
    Param()
    
    Try {
        Return [System.Globalization.CultureInfo]::CurrentCulture.Name
    }
    Catch {
        Return "en-US"
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Test-DMVPNConnection',
    'Test-DMRetailUser',
    'Test-DMRetailComputer',
    'Test-DMRetailUserGroup',
    'Test-DMTerminalSession',
    'Test-DMSharedVDI',
    'Test-DMVirtualMachine',
    'Test-DMServerOS',
    'Get-DMPasswordChangeHotkey',
    'Get-DMSystemLocale'
)

