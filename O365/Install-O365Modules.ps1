<#
.SYNOPSIS
    Check for O365 user and credential files, Test Connect to O365
.DESCRIPTION
    Check for O365 user and credential files:
        o user.txt - Contains O365 UPN (Optional)
        o cred.txt - Contains encrypted O365 password (Required)

    Execute PoSH connect to validate credentials
    Installs missing modules for:  Active Directory Online, Lync Online, SharePointOnline, ExchangeOnline and Office 365 for an organization.
.PARAMETER Path
    Enter alternate path to save files to, defualt is users local app data
.EXAMPLE
    Test for encrypted credential files, load all Office 365 products
    
    Get-O365EncryptedCredentials.ps1
.EXAMPLE
    Test for encrypted credential files, skip loading Skype for Business
    
    Get-O365EncryptedCredentials.ps1 -skipSkype
.NOTES
    Created by Chris Lee
    Date April 20, 2017

    Some code pulled from:
        PoShConnectToOffice365.ps1
        Created by DJacobs for HBS.NET
.LINK
    GitHub: 
    Blogger: 
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

## Install O365 Components
    ##Install MSOnline (Azure Active Directory)
        Write-Verbose -Message "Checking for MSOnline (Office 365) Module"
        ## Office 365
            Try 
                {
                    Write-Verbose -Message "Not connected to MSOnline (Azure Active Directory)"
                    Write-Verbose -Message "$(Get-Date -f o) Importing Module MSOline"
                    Import-Module MSOnline -DisableNameChecking -ErrorAction Stop
                }
            Catch 
                {
                    Write-Verbose -Message "MSOnline Module not found."
                    Write-Verbose -Message 'Check if PowerShell session running in "run as administrator"'
                    If (((whoami /all | select-string S-1-16-12288) -ne $null) -eq $false)
                        {
                            Write-Error 'PowerShell must be ran in "run as administrator to install MSOnline module"'
                            Exit
                        }
                    else 
                        {
                            Write-Host -ForegroundColor Yellow "Installing MSOnline Module"
                            Install-Module MSOnline
                            Try 
                                {
                                    Write-Verbose -Message "$(Get-Date -f o) Importing Module MSOline"
                                    Import-Module MSOnline -DisableNameChecking -ErrorAction Stop
                                }
                            Catch 
                                {
                                    Write-Error -Message "Error. Cannot import module 'MSOnline' because $_" -ErrorAction Stop
                                }
                        }
                    }
            
    ##Install SharePoint Online
        Write-Verbose -Message "Checking for SharePoint Online Module"
        ##SharePoint
                    Try 
                        {                   
                            Write-Verbose -Message "$(Get-Date -f o) Importing SharePoint Module"
                            Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
                        }
                    Catch 
                        {
                            Write-Verbose -Message "SharePoint Online Module not found."
                            $temp = Read-Host "Launch browser to download SharePoint Online Management Shell? [Y]/N"
                            If ($temp -eq $null -OR $temp -eq "" -OR $temp -eq "Y")
                                {
                                    Start-Process https://www.microsoft.com/en-us/download/details.aspx?id=35588      
                                }
                            Else
                                {
                                    Write-Error -Message "Error. Failed to import the module 'Microsoft.Online.SharePoint.PowerShell' because $_" -ErrorAction Stop
                                }
                        }
                        
    ##Install Skype for Business
        Write-Verbose -Message "Checking for Skype for Business Module"
        Try 
            {
                Write-Verbose -Message "$(Get-Date -f o) Importing Module Skype for Business Online Connector"
                Import-Module SkypeOnlineConnector -DisableNameChecking -ErrorAction Stop
            }
        Catch 
            {
                Write-Verbose -Message "SkypeOnlineConnector Module not found."
                $temp = Read-Host "Launch browser to download Skype for Business Online, Windows PowerShell Module? [Y]/N"
                    If ($temp -eq $null -OR $temp -eq "" -OR $temp -eq "Y")
                        {
                            Start-Process https://www.microsoft.com/en-us/download/details.aspx?id=39366    
                        }
                    Else
                        {
                            Write-Error -Message "Error. Failed to import the module 'SkypeOnlineConnector' because $_" -ErrorAction Stop
                        }
            }