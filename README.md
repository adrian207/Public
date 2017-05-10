# Public
Public sharing of scripts created in free my time and for needs in career.

I will attempt to do complete Get-Help Syntax for all scripts posted.
I try to ensure credit given for any code taken from others and request same be followed for my code.

Current Repo structure (Note if you make changes and a script calls on an other it may break the script)

ActiveDirectory

GAFE (Google Apps for Education (Now Google Suit for Education))

O365 (Office365)
    
    Get-O365EncryptedCredentials.ps1
        Check for O365 user and credential files:
            o user.txt - Contains O365 UPN
            o cred.txt - Contains encrypted O365 password

        Load saved Office 365 Credentials from Set-O365EncryptedCredentials.ps1 and either pass or test present

    Install-O365Modules.ps1
        Check for O365 user and credential files:
            o user.txt - Contains O365 UPN (Optional)
            o cred.txt - Contains encrypted O365 password (Required)

        Execute PoSH connect to validate credentials
        Installs missing modules for:  Active Directory Online, Lync Online, SharePointOnline, ExchangeOnline and Office 365 for an organization.


    Invoke-O365AzureSync.ps1
        Run Azure (DirSync) on dedicated remote server

    New-O365PSSession.ps1
        Connect to O365
            If Encrypeted path available will use that otherwise prompts for credentials.

    Remove-O365PSSession.ps1
        Remove open O365 PSSessions

    Set-O365EncryptedCredentials.ps1
        Create following files for use in auto login for O365:
            o user.txt - Contains O365 UPN
            o cred.txt - Contains encrypted O365 password
    
    Test-O365EncryptedCredentials.ps1
        Load Credentials via Get-O365EncryptedCredentials.ps1

        Execute MSOnline (Active Directory Online) connection to validate credentials .
        Will prompt to close when complete to close MSOnline session