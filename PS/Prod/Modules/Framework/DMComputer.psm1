<#
.SYNOPSIS
    Desktop Management Computer Information Module
    
.DESCRIPTION
    Collects comprehensive computer information including AD properties, groups, site, IP addresses.
    Replacement for VBScript ComputerObject class.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: VB/Source/Main/Computer.vbs
#>

# Import required modules
Using Module .\DMCommon.psm1
Using Module .\DMLogger.psm1

<#
.SYNOPSIS
    Gets comprehensive computer information.
    
.DESCRIPTION
    Collects all computer properties needed for Desktop Management operations.
    Returns PSCustomObject with computer details.
    
.OUTPUTS
    PSCustomObject with computer properties
    
.EXAMPLE
    $Computer = Get-DMComputerInfo
    Write-Host "Computer: $($Computer.Name) in site $($Computer.Site)"
#>
Function Get-DMComputerInfo {
    [CmdletBinding()]
    Param()
    
    Try {
        # Get basic computer name
        [String]$Name = $env:COMPUTERNAME
        
        # Get AD information using .NET DirectoryServices
        [Object]$ADInfo = Get-DMComputerADInfo
        
        # Check if we're domain-joined using robust detection
        If (-not $ADInfo.IsDomainJoined) {
            Write-Verbose "Computer is not domain-joined or AD information unavailable"
            
            # Return minimal info for non-domain computers
            Return [PSCustomObject]@{
                PSTypeName = 'DM.ComputerInfo'
                Name = $Name
                DistinguishedName = ""
                Domain = $env:USERDOMAIN
                ShortDomain = $env:USERDOMAIN
                Site = ""
                Groups = @()
                CityCode = "unknown"
                OUMapping = ""
                IPAddresses = Get-DMIPAddresses
                OSCaption = (Get-DMOSInfo).Caption
                IsDesktop = (Get-DMOSInfo).IsDesktop
                IsServer = (Get-DMOSInfo).IsServer
                DomainRole = (Get-DMOSInfo).DomainRole
                IsVPNConnected = Test-DMVPNConnection
            }
        }
        
        # Get group memberships (only if we have a valid DN)
        [Array]$Groups = @()
        If (-not [String]::IsNullOrEmpty($ADInfo.DistinguishedName)) {
            $Groups = Get-DMComputerGroups -DistinguishedName $ADInfo.DistinguishedName
        }
        
        # Extract city code from DN (only if we have a valid DN)
        [String]$CityCode = "unknown"
        If (-not [String]::IsNullOrEmpty($ADInfo.DistinguishedName)) {
            $CityCode = Get-DMCityCode -DistinguishedName $ADInfo.DistinguishedName -IsComputer $True
        }
        
        # Get OU mapping (only if we have a valid DN)
        [String]$OUMapping = ""
        If (-not [String]::IsNullOrEmpty($ADInfo.DistinguishedName)) {
            $OUMapping = Get-DMOUPath -DistinguishedName $ADInfo.DistinguishedName
        }
        
        # Get IP addresses
        [Array]$IPAddresses = Get-DMIPAddresses
        
        # Get OS information
        [Object]$OSInfo = Get-DMOSInfo
        
        # Check if VPN connected
        [Boolean]$IsVPNConnected = Test-DMVPNConnection
        
        # Build computer info object
        [PSCustomObject]$ComputerInfo = [PSCustomObject]@{
            PSTypeName = 'DM.ComputerInfo'
            Name = $Name
            DistinguishedName = $ADInfo.DistinguishedName
            Domain = $ADInfo.DomainDNS
            ShortDomain = $ADInfo.DomainShort
            Site = $ADInfo.Site
            Groups = $Groups
            CityCode = $CityCode
            OUMapping = $OUMapping
            IPAddresses = $IPAddresses
            OSCaption = $OSInfo.Caption
            IsDesktop = $OSInfo.IsDesktop
            IsServer = $OSInfo.IsServer
            DomainRole = $OSInfo.DomainRole
            IsVPNConnected = $IsVPNConnected
        }
        
        Return $ComputerInfo
    }
    Catch {
        Write-Error "Failed to get computer information: $($_.Exception.Message)"
        Return $Null
    }
}

<#
.SYNOPSIS
    Gets Active Directory information for the computer.
    
.DESCRIPTION
    Uses .NET DirectoryServices to retrieve AD properties.
    
.OUTPUTS
    PSCustomObject with DN, Domain, Site
    
.EXAMPLE
    $ADInfo = Get-DMComputerADInfo
#>
Function Get-DMComputerADInfo {
    [CmdletBinding()]
    Param()
    
    Try {
        # Load .NET DirectoryServices
        Add-Type -AssemblyName System.DirectoryServices
        
        # Get current domain information
        [Object]$Domain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
        [String]$DomainDNS = $Domain.Name
        [String]$DomainShort = $Domain.Name.Split('.')[0]
        
        # Create DirectorySearcher for computer
        [Object]$Searcher = New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.Filter = "(&(objectClass=computer)(name=$env:COMPUTERNAME))"
        [Void]$Searcher.PropertiesToLoad.AddRange(@("distinguishedName", "location", "description"))
        
        [Object]$ComputerResult = $Searcher.FindOne()
        
        If ($ComputerResult) {
            [String]$ComputerDN = $ComputerResult.Properties["distinguishedName"][0]
            
            # Get site information from registry (DynamicSiteName)
            [String]$Site = ""
            Try {
                [String]$SiteRegPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Netlogon\Parameters"
                [String]$DynamicSiteName = (Get-ItemProperty -Path $SiteRegPath -Name "DynamicSiteName" -ErrorAction SilentlyContinue).DynamicSiteName
                
                If ($DynamicSiteName) {
                    $Site = $DynamicSiteName
                    Write-DMLog "Retrieved site name from registry: '$Site'" -Level Verbose
                } Else {
                    Write-DMLog "DynamicSiteName registry value not found or empty" -Level Verbose
                }
            } Catch {
                Write-DMLog "Failed to read DynamicSiteName from registry: $($_.Exception.Message)" -Level Verbose
            }
            
            Return [PSCustomObject]@{
                DistinguishedName = $ComputerDN
                DomainDNS = $DomainDNS
                DomainShort = $DomainShort
                Site = $Site
                IsDomainJoined = $True
            }
        } Else {
            Write-Verbose "Computer not found in Active Directory"
            Return [PSCustomObject]@{
                DistinguishedName = ""
                DomainDNS = $DomainDNS
                DomainShort = $DomainShort
                Site = ""
                IsDomainJoined = $False
            }
        }
    }
    Catch {
        Write-Verbose "Failed to get AD computer information using .NET DirectoryServices: $($_.Exception.Message)"
        
        # Fallback to basic domain detection
        Try {
            [Object]$ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
            If ($ComputerSystem -and $ComputerSystem.PartOfDomain) {
                Return [PSCustomObject]@{
                    DistinguishedName = ""
                    DomainDNS = $ComputerSystem.Domain
                    DomainShort = $ComputerSystem.Domain.Split('.')[0]
                    Site = ""
                    IsDomainJoined = $True
                }
            }
        } Catch {
            Write-Verbose "Fallback domain detection also failed: $($_.Exception.Message)"
        }
        
        # Final fallback to environment variables
        [String]$DomainDNS = $env:USERDNSDOMAIN
        [String]$DomainShort = $env:USERDOMAIN
        [Boolean]$IsDomainJoined = Test-DMComputerDomainJoined -ComputerDN "" -DomainDNS $DomainDNS -DomainShort $DomainShort
        
        Return [PSCustomObject]@{
            DistinguishedName = ""
            DomainDNS = $DomainDNS
            DomainShort = $DomainShort
            Site = ""
            IsDomainJoined = $IsDomainJoined
        }
    }
}

<#
.SYNOPSIS
    Tests if computer is domain-joined using multiple detection methods.
    
.DESCRIPTION
    Uses multiple methods to determine if computer is domain-joined:
    1. .NET DirectoryServices results
    2. Environment variables
    3. WMI ComputerSystem.DomainRole
    4. Registry domain information
    
.PARAMETER ComputerDN
    Computer Distinguished Name from .NET DirectoryServices
    
.PARAMETER DomainDNS
    Domain DNS name from .NET DirectoryServices or environment
    
.PARAMETER DomainShort
    Domain short name from .NET DirectoryServices or environment
    
.OUTPUTS
    Boolean - true if domain-joined
    
.EXAMPLE
    $IsDomainJoined = Test-DMComputerDomainJoined -ComputerDN $DN -DomainDNS $DomainDNS -DomainShort $DomainShort
#>
Function Test-DMComputerDomainJoined {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [String]$ComputerDN,
        
        [Parameter(Mandatory=$False)]
        [String]$DomainDNS,
        
        [Parameter(Mandatory=$False)]
        [String]$DomainShort
    )
    
    Try {
        # Method 1: Check if we have valid .NET DirectoryServices results
        If (-not [String]::IsNullOrEmpty($ComputerDN) -and -not [String]::IsNullOrEmpty($DomainDNS)) {
            Return $True
        }
        
        # Method 2: Check environment variables for domain info
        If (-not [String]::IsNullOrEmpty($env:USERDNSDOMAIN) -and $env:USERDNSDOMAIN -ne $env:COMPUTERNAME) {
            Return $True
        }
        
        # Method 3: Check WMI ComputerSystem.DomainRole
        [Object]$ComputerSystem = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        If ($ComputerSystem -and $ComputerSystem.DomainRole -ge 2) {
            Return $True
        }
        
        # Method 4: Check if LOGONSERVER is a domain controller (not local machine)
        If (-not [String]::IsNullOrEmpty($env:LOGONSERVER) -and $env:LOGONSERVER -ne "\\$($env:COMPUTERNAME)") {
            Return $True
        }
        
        Return $False
    }
    Catch {
        Write-Verbose "Error checking domain membership: $($_.Exception.Message)"
        Return $False
    }
}

<#
.SYNOPSIS
    Gets computer group memberships across the forest.
    
.DESCRIPTION
    Queries all domains in the forest for groups containing the computer.
    Returns array of group objects with DN and Name.
    
.PARAMETER DistinguishedName
    Computer's distinguished name
    
.OUTPUTS
    Array of PSCustomObject with GroupDN and GroupName properties
    
.EXAMPLE
    $Groups = Get-DMComputerGroups -DistinguishedName $Computer.DistinguishedName
#>
Function Get-DMComputerGroups {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DistinguishedName
    )
    
    If ([String]::IsNullOrEmpty($DistinguishedName)) {
        Return @()
    }
    
    Try {
        # Load .NET DirectoryServices
        Add-Type -AssemblyName System.DirectoryServices
        
        [Array]$AllGroups = @()
        
        # Get all domains in the forest (like VBScript GetTrees())
        Try {
            [Object]$Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
            [Array]$Domains = $Forest.Domains
            Write-DMLog "Found $($Domains.Count) domains in forest" -Level Verbose
        }
        Catch {
            # Fallback to current domain only
            [Object]$CurrentDomain = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
            [Array]$Domains = @($CurrentDomain)
            Write-DMLog "Using current domain only: $($CurrentDomain.Name)" -Level Verbose
        }
        
        # Query each domain for groups (like VBScript For Each strDomain in arrForrestDomains)
        ForEach ($Domain in $Domains) {
            Try {
                Write-DMLog "Querying computer groups in domain: $($Domain.Name)" -Level Verbose
                
                # Create DirectorySearcher for this domain
                [Object]$Searcher = New-Object System.DirectoryServices.DirectorySearcher
                $Searcher.Filter = "(&(objectClass=group)(member=$DistinguishedName))"
                $Searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
                $Searcher.PageSize = 1000
                [Void]$Searcher.PropertiesToLoad.AddRange(@("distinguishedName", "name"))
                
                # Set search root to domain
                [Object]$DomainEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$($Domain.Name)")
                $Searcher.SearchRoot = $DomainEntry
                
                [Object]$GroupResults = $Searcher.FindAll()
                
                ForEach ($GroupResult in $GroupResults) {
                    [String]$GroupDN = $GroupResult.Properties["distinguishedName"][0]
                    [String]$GroupName = $GroupResult.Properties["name"][0]
                    
                    $AllGroups += [PSCustomObject]@{
                        PSTypeName = 'DM.GroupInfo'
                        DistinguishedName = $GroupDN
                        Name = $GroupName
                    }
                }
                
                Write-DMLog "Found $($GroupResults.Count) computer groups in domain $($Domain.Name)" -Level Verbose
            }
            Catch {
                Write-DMLog "Failed to query computer groups in domain $($Domain.Name): $($_.Exception.Message)" -Level Verbose
            }
        }
        
        Write-DMLog "Total computer groups found across all domains: $($AllGroups.Count)" -Level Verbose
        Return $AllGroups
    }
    Catch {
        Write-DMLog "Failed to get computer groups: $($_.Exception.Message)" -Level Verbose
        Return @()
    }
}

<#
.SYNOPSIS
    Extracts city code from distinguished name.
    
.DESCRIPTION
    Parses DN to extract city code from OU structure.
    Patterns: OU=DEVICES,OU=RESOURCES,OU=<CITY> or OU=USERS,OU=RESOURCES,OU=<CITY>
    Special handling for UAT: OU=RESOURCESUAT
    
.PARAMETER DistinguishedName
    LDAP DN to parse
    
.PARAMETER IsComputer
    True for computer DN, False for user DN
    
.OUTPUTS
    String - city code or "unknown"
    
.EXAMPLE
    $CityCode = Get-DMCityCode -DistinguishedName $DN -IsComputer $True
#>
Function Get-DMCityCode {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DistinguishedName,
        
        [Parameter(Mandatory=$False)]
        [Boolean]$IsComputer = $True
    )
    
    If ([String]::IsNullOrEmpty($DistinguishedName)) {
        Return "unknown"
    }
    
    Try {
        [String]$DNUpper = $DistinguishedName.ToUpper()
        [Array]$Parts = $DistinguishedName -split ','
        
        # Determine OU type based on Computer or User
        [String]$OUType = If ($IsComputer) { "OU=DEVICES" } Else { "OU=USERS" }
        
        # Check for UAT environment first
        If ($DNUpper -match "OU=DEVICES,OU=RESOURCESUAT,OU=" -or $DNUpper -match "OU=USERS,OU=RESOURCESUAT,OU=") {
            For ([Int]$i = 0; $i -lt $Parts.Length; $i++) {
                If ($Parts[$i] -like "*OU=RESOURCESUAT*") {
                    If ($i + 1 -lt $Parts.Length) {
                        [String]$CityOU = $Parts[$i + 1]
                        Return $CityOU.Replace("OU=", "").Replace("ou=", "").Trim()
                    }
                }
            }
        }
        # Check for standard RESOURCES pattern
        ElseIf ($DNUpper -match "$OUType,OU=RESOURCES,OU=") {
            For ([Int]$i = 0; $i -lt $Parts.Length; $i++) {
                If ($Parts[$i] -like "*OU=RESOURCES*" -and $Parts[$i] -notlike "*RESOURCESUAT*") {
                    If ($i + 1 -lt $Parts.Length) {
                        [String]$CityOU = $Parts[$i + 1]
                        Return $CityOU.Replace("OU=", "").Replace("ou=", "").Trim()
                    }
                }
            }
        }
        
        Return "unknown"
    }
    Catch {
        Return "unknown"
    }
}

<#
.SYNOPSIS
    Gets the full OU path in canonical format.
    
.DESCRIPTION
    Converts DN to readable OU path format.
    
.PARAMETER DistinguishedName
    LDAP DN
    
.OUTPUTS
    String - canonical OU path
    
.EXAMPLE
    $OUPath = Get-DMOUPath -DistinguishedName $DN
#>
Function Get-DMOUPath {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DistinguishedName
    )
    
    If ([String]::IsNullOrEmpty($DistinguishedName)) {
        Return ""
    }
    
    Try {
        # Extract DC (domain) and OU components
        [Array]$Parts = $DistinguishedName -split ','
        [Array]$DCParts = @()
        [Array]$OUParts = @()
        
        ForEach ($Part in $Parts) {
            If ($Part -like "DC=*") {
                $DCParts += $Part.Replace("DC=", "").Replace("dc=", "").Trim()
            } ElseIf ($Part -like "OU=*") {
                $OUParts += $Part.Replace("OU=", "").Replace("ou=", "").Trim()
            }
        }
        
        # Build domain path (e.g., "MYMSNGROUP.COM")
        [String]$DomainPath = $DCParts -join '.'
        
        # Reverse OU parts for canonical order
        [Array]::Reverse($OUParts)
        
        # Combine domain + OU path
        If ($OUParts.Count -gt 0) {
            Return "$DomainPath/$($OUParts -join '/')"
        } Else {
            Return $DomainPath
        }
    }
    Catch {
        Return ""
    }
}

<#
.SYNOPSIS
    Gets all IP addresses for the computer.
    
.DESCRIPTION
    Retrieves IP addresses from all enabled network adapters.
    
.OUTPUTS
    Array of IP address strings
    
.EXAMPLE
    $IPs = Get-DMIPAddresses
#>
Function Get-DMIPAddresses {
    [CmdletBinding()]
    Param()
    
    Try {
        [Array]$IPAddresses = @()
        
        # Use CIM instead of WMI for better performance
        [Array]$Configs = Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter "IPEnabled = TRUE"
        
        ForEach ($Config in $Configs) {
            If ($Null -ne $Config.IPAddress) {
                ForEach ($IP in $Config.IPAddress) {
                    $IPAddresses += $IP
                }
            }
        }
        
        Return $IPAddresses
    }
    Catch {
        Return @()
    }
}

<#
.SYNOPSIS
    Gets operating system information.
    
.DESCRIPTION
    Retrieves OS caption, determines if desktop or server.
    
.OUTPUTS
    PSCustomObject with OS properties
    
.EXAMPLE
    $OSInfo = Get-DMOSInfo
#>
Function Get-DMOSInfo {
    [CmdletBinding()]
    Param()
    
    Try {
        [Object]$OS = Get-CimInstance -ClassName Win32_OperatingSystem
        [Object]$Computer = Get-CimInstance -ClassName Win32_ComputerSystem
        
        [String]$Caption = $OS.Caption
        [Boolean]$IsServer = $Caption -like "*Server*" -or $Computer.DomainRole -ge 3
        [Boolean]$IsDesktop = -not $IsServer
        
        Return [PSCustomObject]@{
            PSTypeName = 'DM.OSInfo'
            Caption = $Caption
            IsDesktop = $IsDesktop
            IsServer = $IsServer
            DomainRole = $Computer.DomainRole
        }
    }
    Catch {
        Return [PSCustomObject]@{
            PSTypeName = 'DM.OSInfo'
            Caption = "Unknown"
            IsDesktop = $True
            IsServer = $False
            DomainRole = 0
        }
    }
}

<#
.SYNOPSIS
    Tests if computer is member of a specific group.
    
.DESCRIPTION
    Checks if computer is member of a group (supports nested groups).
    
.PARAMETER ComputerInfo
    Computer info object from Get-DMComputerInfo
    
.PARAMETER GroupName
    Group name to check
    
.OUTPUTS
    Boolean - true if member
    
.EXAMPLE
    $IsMember = Test-DMComputerGroupMembership -ComputerInfo $Computer -GroupName "Pilot Desktop Management Script"
#>
Function Test-DMComputerGroupMembership {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$ComputerInfo,
        
        [Parameter(Mandatory=$True)]
        [String]$GroupName
    )
    
    If ($Null -eq $ComputerInfo.Groups -or $ComputerInfo.Groups.Count -eq 0) {
        Return $False
    }
    
    ForEach ($Group in $ComputerInfo.Groups) {
        If ($Group.GroupName -eq $GroupName -or $Group.GroupName -like "*$GroupName*") {
            Return $True
        }
    }
    
    Return $False
}

# Export module members
Export-ModuleMember -Function @(
    'Get-DMComputerInfo',
    'Get-DMComputerADInfo',
    'Test-DMComputerDomainJoined',
    'Get-DMComputerGroups',
    'Get-DMCityCode',
    'Get-DMOUPath',
    'Get-DMIPAddresses',
    'Get-DMOSInfo',
    'Test-DMComputerGroupMembership'
)

