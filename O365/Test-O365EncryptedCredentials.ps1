<#
.SYNOPSIS
    Test Connect to Office 365 using saved Office 365 Credentials from Set-O365EncryptedCredentials.ps1
.DESCRIPTION
    Load Credentials via Get-O365EncryptedCredentials.ps1

    Execute MSOnline (Active Directory Online) connection to validate credentials .
    Will prompt to close when complete to close MSOnline session
.PARAMETER Path
    Enter alternate path to save files to, defualt is users local app data
.EXAMPLE
    Test Saved O365 credentials are valid.

    Test-O365EncryptedCredentials.ps1
.NOTES
    Created by Chris Lee
    Date April 20, 2017

    Some code pulled from:
        PoShConnectToOffice365.ps1
        Created by DJacobs for HBS.NET
.LINK
    GitHub: https://github.com/clee1107/Public/blob/master/O365/Test-O365EncryptedCredentials.ps1
    Blogger: http://www.myitresourcebook.com/2017/05/test-o365encryptedcredentialsps1.html
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

## Get Credentials
    Write-Verbose -Message "Calling Get-O365EnccryptedCredentials.ps1 to load saved O365 credentials."
    $Cred = Get-O365EncryptedCredentials.ps1 -Path $Path
    If ($Cred.username -eq $null -OR $Cred.username -eq "")
        {
            Write-Host -ForegroundColor Red "Failed to load Encrypted Credentials"
            Write-Error -Message "Failed to load Encrypted Credentials"
            Exit
        }
    Else
        {
            Write-Host -ForegroundColor Green "Encrypted Credentials loaded"
            Write-Verbose -Message "Received from Get-O365EncryptedCredentials:"
            Write-Verbose -Message "Username: $($Cred.UserName)"
        }

## Connect of O365 with error checks
    ##Connect MSOnline (Azure Active Directory)
        Write-Verbose -Message "Testing MSOnline (Office 365) Connection with Credentials"
        ## Office 365
            Try 
                {
                    Write-Verbose -Message "$(Get-Date -f o) Importing Module MSOline"
                    Import-Module -Name MSOnline -DisableNameChecking -ErrorAction Stop
                    Write-Verbose -Message "$(Get-Date -f o) Imported Module MSOline"
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
            Try 
                {
                    Write-Verbose -Message "Connecting MSOnline"
                    Connect-MsolService -Credential $cred -ErrorAction Stop
                    Get-MsolDomain -ErrorAction Stop > $null
                    Write-Host -ForegroundColor Green "Connected to MSOnline (Azure Active Directory)"
                    Write-Verbose -Message "Connected MSOnline"
                }
            Catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException] 
                {
                    Write-Error -Message "Error. Failed to Connect to MSOnline because bad username or password." -ErrorAction Stop
                }
            Catch 
                {
                    Write-Error -Message "Error. Failed to connect to MSOnline because $_" -ErrorAction Stop
                }

##Close open O365 Sessions
    $ExitPowershell = Read-Host -Prompt "Disconnect from O365 (Will close current Powershell Window) [Y]/N"
    If ($ExitPowershell -eq "Y" -OR $ExitPowershell -eq $null -OR $ExitPowershell -eq "")
        {
            stop-process -Id $PID
        }