# Public
Public shares of scripts created in free time and for needs in career.

I will attempt to do complete Get-Help Syntax for all scripts posted.
I try to ensure credit given for any code taken from others and request same be followed for my code.

Currently scripts are broken into target types:

ActiveDirectory

GAFE (Google Apps for Education (Now Google Suit for Education))

O365 (Office365)
    
    Get-O365EncryptedCredentials.ps1
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
            o user.txt - Contains O365 UPN (Optional)
            o cred.txt - Contains encrypted O365 password (Required)
