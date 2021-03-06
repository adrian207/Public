<#
.SYNOPSIS
    Load saved Office 365 Credentials from Set-O365EncryptedCredentials.ps1 and either pass or test present
.DESCRIPTION
    Check for O365 user and credential files:
        o user.txt - Contains O365 UPN
        o cred.txt - Contains encrypted O365 password

    Load saved Office 365 Credentials from Set-O365EncryptedCredentials.ps1 and either pass or test present
.PARAMETER Test
    $fale to operate in "Pass-thru Mode" (Default)
    $true to validate user name and password loaded from saved credentials files
.PARAMETER Path
    Enter alternate path to save files to, defualt is users local app data
.EXAMPLE
    Test for encrypted credential files
    
    Get-O365EncryptedCredentials.ps1 -Test
.EXAMPLE
    Pass encrypted credential to calling script
    
    Get-O365EncryptedCredentials.ps1
.EXAMPLE
    Test for encrypted credential files with verbose messages
    
    Get-O365EncryptedCredentials.ps1 -Test -Verbose
.NOTES
    Created by Chris Lee
    Date May 9th, 2017
.LINK
    GitHub: https://github.com/clee1107/Public/blob/master/O365/Get-O365EncryptedCredentials.ps1
    Blogger: http://www.myitresourcebook.com/2017/05/get-o365encryptedcredentialsps1_9.html
#>

[Cmdletbinding()]
Param
(
    [switch]
    $Test,
    [String]
    $Path = [Environment]::GetFolderPath("LocalApplicationData")
)

#################################
## DO NOT EDIT BELOW THIS LINE ##
#################################

##Load User name
    Write-Verbose -Message "Checking for user file"
    IF (!(Test-Path -Path "$Path\O365user.txt"))
        {
            Write-Host -ForegroundColor Red "No user file found"
            Write-Error -Message "No user file found"
            Exit
        }
    else
        {
            Write-Host -ForegroundColor Green "User File found"
            $AdminName = Get-Content "$Path\O365user.txt"
        }

##Load Encrypted password
    Write-Verbose -Message "Checking for required password file"
    IF (!(Test-Path -Path "$Path\O365cred.txt"))
        {
            Write-Host -ForegroundColor Red "No password file found"
            Write-Error -Message "No password file found"
            Exit
        }
    else
        {
            Write-Host -ForegroundColor Green "Password File Found"
            $Pass = Get-Content "$Path\O365cred.txt" | ConvertTo-SecureString
        }

## Create Cred variable
    Write-Verbose -Message "Creating Credential variable from files"
    $Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass

## Check if testing or passing credentials
    Write-Verbose -Message "Check if in test mode or pass-thru mode"
    If ($test)
        {
            Write-Verbose -Message "Test mode"
            ## Display loaded credentias
                Write-Host "Username: $($Cred.UserName)"
                Write-Host "Password: $($Cred.Password)"
        }
    Else 
        {
            Write-Verbose -Message "Pass-thru mode"
            ## Passing Cred variable to other script
                $Cred
        }