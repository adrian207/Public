<#
.SYNOPSIS
    Create User and Password files for future use
.DESCRIPTION
    Create following files for use in auto login for O365:
        o user.txt - Contains O365 UPN (Optional)
        o cred.txt - Contains encrypted O365 password (Required)
.PARAMETER Path
    Enter alternate path to save files to, defualt is users local app data
.EXAMPLE
    Set-O365EncryptedCredentials.ps1
.NOTES
    Created by Chris Lee
    Date April 20, 2017
.LINK   
#>

[Cmdletbinding()]
Param
(
    [String]
    $Path = [Environment]::GetFolderPath("LocalApplicationData")
)

#################################
## DO NOT EDIT BELOW THIS LINE ##
#################################

##Create User account if provided
    Read-Host -Prompt "Enter your tenant UPN" `
        | Out-File "$Path\O365user.txt"

##Create Password
    Read-Host -Prompt "Enter your tenant password" -AsSecureString `
        | ConvertFrom-SecureString `
        | Out-File "$Path\O365cred.txt"
