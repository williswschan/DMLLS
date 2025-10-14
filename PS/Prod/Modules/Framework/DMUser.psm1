<#
.SYNOPSIS
    Desktop Management User Information Module
    
.DESCRIPTION
    Collects comprehensive user information including AD properties, groups, logon server.
    Replacement for VBScript UserObject class.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: VB/Source/Main/User.vbs
#>

# Import required modules
Using Module .\DMCommon.psm1
Using Module .\DMLogger.psm1

<#
.SYNOPSIS
    Gets comprehensive user information.
    
.DESCRIPTION
    Collects all user properties needed for Desktop Management operations.
    Returns PSCustomObject with user details.
    
.OUTPUTS
    PSCustomObject with user properties
    
.EXAMPLE
    $User = Get-DMUserInfo
    Write-Host "User: $($User.Name) in domain $($User.Domain)"
#>
Function Get-DMUserInfo {
    [CmdletBinding()]
    Param()
    
    Try {
        # Get basic user information from environment
        [Hashtable]$EnvInfo = Get-DMUserEnvironmentInfo
        
        # Get AD Distinguished Name
        [String]$DistinguishedName = Get-DMUserDN
        
        # Check if we got valid AD info (domain-joined)
        If ([String]::IsNullOrEmpty($DistinguishedName)) {
            Write-Verbose "User is not logged into a domain or AD information unavailable"
            
            # Return minimal info for non-domain users
            Return [PSCustomObject]@{
                PSTypeName = 'DM.UserInfo'
                Name = $EnvInfo.UserName
                Domain = $EnvInfo.DomainDNS
                ShortDomain = $EnvInfo.DomainShort
                LogonServer = $EnvInfo.LogonServer
                DistinguishedName = ""
                Groups = @()
                CityCode = "unknown"
                OUMapping = ""
                IsTerminalSession = (Test-DMTerminalSession).IsTerminalSession
                SessionType = (Test-DMTerminalSession).SessionType
                SessionName = (Test-DMTerminalSession).SessionName
            }
        }
        
        # Get group memberships (only if we have a valid DN)
        [Array]$Groups = @()
        If (-not [String]::IsNullOrEmpty($DistinguishedName)) {
            $Groups = Get-DMUserGroups -DistinguishedName $DistinguishedName
        }
        
        # Extract city code from DN (only if we have a valid DN)
        [String]$CityCode = "unknown"
        If (-not [String]::IsNullOrEmpty($DistinguishedName)) {
            $CityCode = Get-DMCityCode -DistinguishedName $DistinguishedName -IsComputer $False
        }
        
        # Get OU mapping (only if we have a valid DN)
        [String]$OUMapping = ""
        If (-not [String]::IsNullOrEmpty($DistinguishedName)) {
            $OUMapping = Get-DMOUPath -DistinguishedName $DistinguishedName
        }
        
        # Check if terminal session
        [Object]$SessionInfo = Test-DMTerminalSession
        
        # Build user info object
        [PSCustomObject]$UserInfo = [PSCustomObject]@{
            PSTypeName = 'DM.UserInfo'
            Name = $EnvInfo.UserName
            Domain = $EnvInfo.DomainDNS
            ShortDomain = $EnvInfo.DomainShort
            LogonServer = $EnvInfo.LogonServer
            DistinguishedName = $DistinguishedName
            Groups = $Groups
            CityCode = $CityCode
            OUMapping = $OUMapping
            IsTerminalSession = $SessionInfo.IsTerminalSession
            SessionType = $SessionInfo.SessionType
            SessionName = $SessionInfo.SessionName
        }
        
        Return $UserInfo
    }
    Catch {
        Write-Error "Failed to get user information: $($_.Exception.Message)"
        Return $Null
    }
}

<#
.SYNOPSIS
    Gets user information from environment variables.
    
.DESCRIPTION
    Retrieves USERNAME, USERDNSDOMAIN, USERDOMAIN, LOGONSERVER from environment.
    
.OUTPUTS
    Hashtable with user environment info
    
.EXAMPLE
    $EnvInfo = Get-DMUserEnvironmentInfo
#>
Function Get-DMUserEnvironmentInfo {
    [CmdletBinding()]
    Param()
    
    Try {
        [String]$UserName = $env:USERNAME
        [String]$DomainDNS = $env:USERDNSDOMAIN
        [String]$DomainShort = $env:USERDOMAIN
        [String]$LogonServer = $env:LOGONSERVER
        
        # Handle null/empty values
        If ([String]::IsNullOrEmpty($UserName)) { $UserName = "" }
        If ([String]::IsNullOrEmpty($DomainDNS)) { $DomainDNS = "" }
        If ([String]::IsNullOrEmpty($DomainShort)) { $DomainShort = "" }
        If ([String]::IsNullOrEmpty($LogonServer)) { $LogonServer = "" }
        
        # Remove leading backslashes from LOGONSERVER
        $LogonServer = $LogonServer.TrimStart('\')
        
        Return @{
            UserName = $UserName
            DomainDNS = $DomainDNS
            DomainShort = $DomainShort
            LogonServer = $LogonServer
        }
    }
    Catch {
        Return @{
            UserName = ""
            DomainDNS = ""
            DomainShort = ""
            LogonServer = ""
        }
    }
}

<#
.SYNOPSIS
    Gets user's Distinguished Name from Active Directory.
    
.DESCRIPTION
    Uses ADSystemInfo COM object to retrieve user DN.
    
.OUTPUTS
    String - user DN
    
.EXAMPLE
    $DN = Get-DMUserDN
#>
Function Get-DMUserDN {
    [CmdletBinding()]
    Param()
    
    Try {
        # Load .NET DirectoryServices
        Add-Type -AssemblyName System.DirectoryServices
        
        # Create DirectorySearcher for user
        [Object]$Searcher = New-Object System.DirectoryServices.DirectorySearcher
        $Searcher.Filter = "(&(objectClass=user)(samAccountName=$env:USERNAME))"
        [Void]$Searcher.PropertiesToLoad.Add("distinguishedName")
        
        [Object]$UserResult = $Searcher.FindOne()
        
        If ($UserResult) {
            [String]$UserDN = $UserResult.Properties["distinguishedName"][0]
            Return $UserDN
        } Else {
            Write-Verbose "User not found in Active Directory"
            Return ""
        }
    }
    Catch {
        Write-Verbose "Failed to get user DN using .NET DirectoryServices: $($_.Exception.Message)"
        Return ""
    }
}

<#
.SYNOPSIS
    Gets user group memberships across the forest.
    
.DESCRIPTION
    Queries all domains in the forest for groups containing the user.
    Returns array of group objects with DN and Name.
    
.PARAMETER DistinguishedName
    User's distinguished name
    
.OUTPUTS
    Array of PSCustomObject with GroupDN and GroupName properties
    
.EXAMPLE
    $Groups = Get-DMUserGroups -DistinguishedName $User.DistinguishedName
#>
Function Get-DMUserGroups {
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
                Write-DMLog "Querying groups in domain: $($Domain.Name)" -Level Verbose
                
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
                
                Write-DMLog "Found $($GroupResults.Count) groups in domain $($Domain.Name)" -Level Verbose
            }
            Catch {
                Write-DMLog "Failed to query groups in domain $($Domain.Name): $($_.Exception.Message)" -Level Verbose
            }
        }
        
        Write-DMLog "Total groups found across all domains: $($AllGroups.Count)" -Level Verbose
        Return $AllGroups
    }
    Catch {
        Write-DMLog "Failed to get user groups: $($_.Exception.Message)" -Level Verbose
        Return @()
    }
}

<#
.SYNOPSIS
    Tests if user is member of a specific group.
    
.DESCRIPTION
    Checks if user is member of a group by name.
    
.PARAMETER UserInfo
    User info object from Get-DMUserInfo
    
.PARAMETER GroupName
    Group name to check
    
.OUTPUTS
    Boolean - true if member
    
.EXAMPLE
    $IsMember = Test-DMUserGroupMembership -UserInfo $User -GroupName "Laptop Offline PC"
#>
Function Test-DMUserGroupMembership {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$True)]
        [String]$GroupName
    )
    
    If ($Null -eq $UserInfo.Groups -or $UserInfo.Groups.Count -eq 0) {
        Return $False
    }
    
    ForEach ($Group in $UserInfo.Groups) {
        If ($Group.GroupName -eq $GroupName -or $Group.GroupName -like "*$GroupName*") {
            Return $True
        }
    }
    
    Return $False
}

<#
.SYNOPSIS
    Gets user's email address from LDAP.
    
.DESCRIPTION
    Queries Active Directory for user's email attribute.
    Used for PST inventory validation.
    
.PARAMETER DistinguishedName
    User's DN
    
.PARAMETER Domain
    Domain to query
    
.OUTPUTS
    String - email address or empty string
    
.EXAMPLE
    $Email = Get-DMUserEmail -DistinguishedName $User.DistinguishedName -Domain $User.Domain
#>
Function Get-DMUserEmail {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DistinguishedName,
        
        [Parameter(Mandatory=$True)]
        [String]$Domain
    )
    
    If ([String]::IsNullOrEmpty($DistinguishedName) -or [String]::IsNullOrEmpty($Domain)) {
        Return ""
    }
    
    Try {
        # Build LDAP query for email attribute
        [String]$LDAPPath = "LDAP://$Domain/$DistinguishedName"
        
        [Object]$User = [ADSI]$LDAPPath
        
        If ($Null -ne $User.mail -and $User.mail.Count -gt 0) {
            Return $User.mail[0].ToString()
        }
        
        Return ""
    }
    Catch {
        Write-Verbose "Failed to get user email: $($_.Exception.Message)"
        Return ""
    }
}

<#
.SYNOPSIS
    Gets user's password expiry information from LDAP.
    
.DESCRIPTION
    Queries AD for password last changed and max age to calculate expiry.
    
.PARAMETER DistinguishedName
    User's DN
    
.PARAMETER Domain
    Domain to query
    
.OUTPUTS
    PSCustomObject with password expiry information
    
.EXAMPLE
    $PwdInfo = Get-DMUserPasswordExpiry -DistinguishedName $User.DistinguishedName -Domain $User.Domain
#>
Function Get-DMUserPasswordExpiry {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$DistinguishedName,
        
        [Parameter(Mandatory=$True)]
        [String]$Domain
    )
    
    Try {
        # Get user object
        [String]$LDAPPath = "LDAP://$Domain/$DistinguishedName"
        Write-Verbose "Attempting to bind to: $LDAPPath"
        
        [Object]$User = [ADSI]$LDAPPath
        
        # Verify user object is valid
        If ($Null -eq $User -or $Null -eq $User.Path -or [String]::IsNullOrEmpty($User.Path)) {
            Write-Verbose "Failed to bind to user object: $LDAPPath"
            Write-Verbose "User object Path: $($User.Path)"
            Throw "Unable to retrieve user object from Active Directory (LDAP bind failed)"
        }
        
        Write-Verbose "Successfully bound to user object: $($User.Path)"
        
        # Get domain object for max password age
        [String]$DomainPath = "LDAP://$Domain"
        [Object]$DomainObj = [ADSI]$DomainPath
        
        # Check if password never expires
        If ($Null -eq $User.userAccountControl -or $User.userAccountControl.Count -eq 0) {
            Write-Verbose "userAccountControl property is not available"
            Throw "Unable to read user account properties"
        }
        
        [Int]$UserFlags = $User.userAccountControl[0]
        [Boolean]$PasswordNeverExpires = ($UserFlags -band 0x10000) -ne 0
        
        If ($PasswordNeverExpires) {
            Return [PSCustomObject]@{
                PSTypeName = 'DM.PasswordExpiryInfo'
                PasswordNeverExpires = $True
                PasswordLastChanged = $Null
                PasswordExpiryDate = $Null
                DaysUntilExpiry = -1
            }
        }
        
        # Get password last changed
        If ($Null -eq $User.pwdLastSet -or $User.pwdLastSet.Count -eq 0) {
            Write-Verbose "pwdLastSet property is not available"
            Throw "Unable to read password last set date"
        }
        
        [Object]$PwdLastSet = $User.ConvertLargeIntegerToInt64($User.pwdLastSet[0])
        [DateTime]$PasswordLastChanged = [DateTime]::FromFileTime($PwdLastSet)
        
        # Get max password age from domain
        If ($Null -eq $DomainObj.maxPwdAge -or $DomainObj.maxPwdAge.Count -eq 0) {
            Write-Verbose "maxPwdAge property is not available"
            Throw "Unable to read domain password policy"
        }
        
        [Object]$MaxPwdAge = $DomainObj.ConvertLargeIntegerToInt64($DomainObj.maxPwdAge[0])
        
        # maxPwdAge is stored as a negative value in AD
        # Check if it's Int64.MinValue or would cause DateTime overflow
        If ($MaxPwdAge -eq [Int64]::MinValue -or $MaxPwdAge -eq 0) {
            # Password never expires at domain level
            Return [PSCustomObject]@{
                PSTypeName = 'DM.PasswordExpiryInfo'
                PasswordNeverExpires = $True
                PasswordLastChanged = $PasswordLastChanged
                PasswordExpiryDate = $Null
                DaysUntilExpiry = -1
            }
        } ElseIf ($MaxPwdAge -lt 0) {
            [TimeSpan]$MaxAge = [TimeSpan]::FromTicks(-$MaxPwdAge)
        } Else {
            [TimeSpan]$MaxAge = [TimeSpan]::FromTicks($MaxPwdAge)
        }
        
        # Calculate expiry date with overflow protection
        Try {
            [DateTime]$ExpiryDate = $PasswordLastChanged.Add($MaxAge)
            [Int]$DaysLeft = ([Int]($ExpiryDate - (Get-Date)).TotalDays)
        } Catch {
            # DateTime overflow - treat as never expires
            Write-Verbose "Password expiry date calculation overflow - treating as never expires"
            Return [PSCustomObject]@{
                PSTypeName = 'DM.PasswordExpiryInfo'
                PasswordNeverExpires = $True
                PasswordLastChanged = $PasswordLastChanged
                PasswordExpiryDate = $Null
                DaysUntilExpiry = -1
            }
        }
        
        Return [PSCustomObject]@{
            PSTypeName = 'DM.PasswordExpiryInfo'
            PasswordNeverExpires = $False
            PasswordLastChanged = $PasswordLastChanged
            PasswordExpiryDate = $ExpiryDate
            DaysUntilExpiry = $DaysLeft
        }
    }
    Catch {
        Write-Warning "Failed to get password expiry info: $($_.Exception.Message)"
        Return [PSCustomObject]@{
            PSTypeName = 'DM.PasswordExpiryInfo'
            PasswordNeverExpires = $False
            PasswordLastChanged = $Null
            PasswordExpiryDate = $Null
            DaysUntilExpiry = -1
        }
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Get-DMUserInfo',
    'Get-DMUserEnvironmentInfo',
    'Get-DMUserDN',
    'Get-DMUserGroups',
    'Test-DMUserGroupMembership',
    'Get-DMUserEmail',
    'Get-DMUserPasswordExpiry'
)

