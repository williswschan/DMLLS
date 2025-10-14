<#
.SYNOPSIS
    Desktop Management Password Expiry Notification Module
    
.DESCRIPTION
    Shows password expiry notifications to users with multi-language support.
    Displays warnings when password is about to expire (within 14 days).
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
    Replacement for: PasswordExpiryNotification.vbs
#>

# Import required modules
Using Module .\Test-Environment.psm1
Using Module ..\Framework\DMLogger.psm1
Using Module ..\Framework\DMUser.psm1

<#
.SYNOPSIS
    Shows password expiry notification if needed.
    
.DESCRIPTION
    Checks password expiry and shows notification if within warning threshold (14 days).
    Supports multi-language (Japanese and English).
    
.PARAMETER UserInfo
    User information object
    
.PARAMETER WarningDays
    Days before expiry to start warning (default: 14)
    
.OUTPUTS
    Boolean - true if successful or no warning needed
    
.EXAMPLE
    Show-DMPasswordExpiryNotification -UserInfo $User
#>
Function Show-DMPasswordExpiryNotification {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [PSCustomObject]$UserInfo,
        
        [Parameter(Mandatory=$False)]
        [Int]$WarningDays = 14
    )
    
    Try {
        Write-DMLog "PasswordExpiryNotification: Starting" -Level Info
        
        # Check if user is domain-joined
        If ([String]::IsNullOrEmpty($UserInfo.DistinguishedName)) {
            Write-DMLog "PasswordExpiryNotification: User is not domain-joined, skipping" -Level Verbose
            Return $True
        }
        
        Write-DMLog "PasswordExpiryNotification: Checking user account properties" -Level Info
        
        # Get password expiry information
        [Object]$PasswordInfo = Get-DMUserPasswordExpiry -DistinguishedName $UserInfo.DistinguishedName -Domain $UserInfo.Domain
        
        # Check if password never expires
        If ($PasswordInfo.PasswordNeverExpires) {
            Write-DMLog "PasswordExpiryNotification: The user account has a non-expiring password. Notification not required" -Level Info
            Write-DMLog "PasswordExpiryNotification: Completed" -Level Info
            Return $True
        }
        
        # Check if we got valid password info
        If ($Null -eq $PasswordInfo -or $Null -eq $PasswordInfo.PasswordExpiryDate) {
            Write-DMLog "PasswordExpiryNotification: Could not retrieve password expiry information, skipping" -Level Verbose
            Return $True
        }
        
        # Check days left
        [Int]$DaysLeft = $PasswordInfo.DaysUntilExpiry
        
        If ($DaysLeft -lt 0) {
            Write-DMLog "PasswordExpiryNotification: Password has already expired" -Level Warning
            $DaysLeft = 0
        }
        
        # Show notification if within warning threshold
        If ($DaysLeft -lt $WarningDays -and $DaysLeft -ge 0) {
            Write-DMLog "PasswordExpiryNotification: '$DaysLeft' day(s) left for the password to expire" -Level Info
            
            # Check if Retail user with Hyper-V VDI (skip notification)
            [Boolean]$IsRetail = Test-DMRetailUser -DistinguishedName $UserInfo.DistinguishedName
            
            If ($IsRetail) {
                Write-DMLog "PasswordExpiryNotification: User belongs to Retail" -Level Info
                
                [Object]$VMInfo = Test-DMVirtualMachine
                If ($VMInfo.Platform -eq "Hyper-V") {
                    Write-DMLog "PasswordExpiryNotification: Retail VDI Computer. Terminate" -Level Info
                    Return $True
                }
                
                Write-DMLog "PasswordExpiryNotification: NOT Retail VDI Computer" -Level Info
            } Else {
                Write-DMLog "PasswordExpiryNotification: The user does not belong to Retail" -Level Info
            }
            
            # Get language
            [String]$Language = Get-DMPasswordNotificationLanguage -IsRetail $IsRetail
            Write-DMLog "PasswordExpiryNotification: Language: '$Language'" -Level Info
            
            # Get session hotkey
            [Object]$SessionInfo = Test-DMTerminalSession
            [String]$Hotkey = Get-DMPasswordChangeHotkey -SessionInfo $SessionInfo
            Write-DMLog "PasswordExpiryNotification: Hotkey: '$Hotkey'" -Level Info
            
            # Build and show notification
            [String]$Message = Get-DMPasswordExpiryMessage -DaysLeft $DaysLeft -ExpiryDate $PasswordInfo.PasswordExpiryDate -Hotkey $Hotkey -Language $Language
            
            Show-DMNotificationPopup -Message $Message -Title "Password Expiry Notification"
            
            Write-DMLog "PasswordExpiryNotification: Notification displayed to user" -Level Info
        } Else {
            Write-DMLog "PasswordExpiryNotification: Password expires in $DaysLeft day(s), no notification needed" -Level Verbose
        }
        
        Write-DMLog "PasswordExpiryNotification: Completed" -Level Info
        Return $True
    }
    Catch {
        Write-DMLog "PasswordExpiryNotification: Error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Gets the appropriate language for password notification.
    
.DESCRIPTION
    Returns language code based on Retail status and system locale.
    
.PARAMETER IsRetail
    Whether user is Retail
    
.OUTPUTS
    String - language code (ja-JP or en-US)
#>
Function Get-DMPasswordNotificationLanguage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Boolean]$IsRetail
    )
    
    [String]$SystemLocale = Get-DMSystemLocale
    
    If ($IsRetail) {
        # Retail defaults to Japanese
        If ([String]::IsNullOrEmpty($SystemLocale)) {
            Return "ja-JP"
        }
        Return $SystemLocale
    } Else {
        # Wholesale defaults to English
        If ([String]::IsNullOrEmpty($SystemLocale)) {
            Return "en-US"
        }
        Return $SystemLocale
    }
}

<#
.SYNOPSIS
    Builds password expiry notification message.
    
.DESCRIPTION
    Creates localized message text based on language.
    
.PARAMETER DaysLeft
    Days until password expires
    
.PARAMETER ExpiryDate
    Password expiry date
    
.PARAMETER Hotkey
    Hotkey combination for password change
    
.PARAMETER Language
    Language code
    
.OUTPUTS
    String - notification message
#>
Function Get-DMPasswordExpiryMessage {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Int]$DaysLeft,
        
        [Parameter(Mandatory=$True)]
        [DateTime]$ExpiryDate,
        
        [Parameter(Mandatory=$True)]
        [String]$Hotkey,
        
        [Parameter(Mandatory=$False)]
        [String]$Language = "en-US"
    )
    
    If ($Language -eq "ja-JP") {
        # Japanese message - using here-string for better encoding compatibility
        [String]$Year = $ExpiryDate.Year
        [String]$Month = $ExpiryDate.Month
        [String]$Day = $ExpiryDate.Day
        [String]$Hour = $ExpiryDate.Hour
        [String]$Minute = $ExpiryDate.Minute
        [String]$Second = $ExpiryDate.Second
        
        # Build message parts separately to avoid encoding issues
        [String]$Part1 = "Your password will expire in $DaysLeft day(s)."
        [String]$Part2 = "(Expiry Date: ${Year}/${Month}/${Day} ${Hour}:${Minute}:${Second})"
        [String]$Part3 = "Please press $Hotkey and select 'Change a password' to change your password."
        
        Return "${Part1}`n`n${Part2}`n`n${Part3}"
    } Else {
        # English message
        [String]$ExpiryDateFormatted = $ExpiryDate.ToString("yyyy-MM-dd HH:mm:ss")
        [String]$Part1 = "Your password will expire in $DaysLeft day(s) at $ExpiryDateFormatted."
        [String]$Part2 = "Please press $Hotkey and select 'Change a password' to change your password."
        
        Return "${Part1}`n`n${Part2}"
    }
}

<#
.SYNOPSIS
    Shows a notification popup to the user.
    
.DESCRIPTION
    Displays a message box with the notification.
    
.PARAMETER Message
    Message to display
    
.PARAMETER Title
    Popup title
    
.OUTPUTS
    Boolean - true if displayed
#>
Function Show-DMNotificationPopup {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$Message,
        
        [Parameter(Mandatory=$False)]
        [String]$Title = "Notification"
    )
    
    Try {
        # Use WScript.Shell Popup method
        [Object]$WshShell = New-Object -ComObject WScript.Shell
        
        # Popup(Text, SecondsToWait, Title, Type)
        # Type: 0 = OK button, 48 = Warning icon
        [Void]$WshShell.Popup($Message, 0, $Title, 48)
        
        Return $True
    }
    Catch {
        Write-DMLog "Show Notification: Error displaying popup: $($_.Exception.Message)" -Level Warning
        Return $False
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Show-DMPasswordExpiryNotification',
    'Get-DMPasswordNotificationLanguage',
    'Get-DMPasswordExpiryMessage',
    'Show-DMNotificationPopup'
)

