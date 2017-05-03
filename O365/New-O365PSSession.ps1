<#
.SYNOPSIS
    Connect to O365
.DESCRIPTION
    Connect to O365
        If Encrypeted path available will use that otherwise prompts for credentials.
.PARAMETER SkipSharePoint
    If no DNS for SharePoint, add this to skip SharePoint options
.PARAMETER SkipSkype
    If no DNS for Lync, add this to skip Lync options
.PARAMETER SkipExchange
    If no DNS for Exchange, add this to skip Exchange options
.PARAMETER SkipSecurity
    If no DNS for Security & Compliance Center, add this to skip Security & Compliance Center options
.PARAMETER Path
    Enter alternate path for ecrypted credential files, defualt is users local app data
.EXAMPLE
    Opens PowerShell Session for all Office365 products
    
    New-O365PSSession.ps1
.EXAMPLE
    Opens PowerShell Session for all Office365 products except Exchange online
    
    New-O365PSSession.ps1 -skipExchange
.NOTES
    Created by Chris Lee
    Date April 20, 2017

    Some code pulled from:
        PoShConnectToOffice365.ps1
        Created by DJacobs for HBS.NET
.LINK   
#>

[Cmdletbinding()]
Param
(
    [switch]
    $SkipSharePoint,
    [switch]
    $SkipSkype,
    [switch]
    $SkipExchange,
    [switch]
    $SkipSecurity,
    [String]
    $Path = [Environment]::GetFolderPath("LocalApplicationData")
)

#################################
## DO NOT EDIT BELOW THIS LINE ##
#################################

## region Functions
    Function Test-Verbose
        {
            [cmdletbinding()]
            Param()
            Write-Verbose "Verbose output"
            "Regular output"
        }
    Test-Verbose
## endregion

## Check for Credential Files present
    ##Check for User file
        Write-Verbose -Message "Checking for optional user file"
        IF (!(Test-Path -Path "$Path\O365user.txt"))
            {
                $AdminName = Read-Host -Prompt "Enter your tenant UPN"       
            }
        else
            {
                Write-Host -ForegroundColor Green "User File found"
                $AdminName = Get-Content "$Path\O365user.txt"
            }
    ##Load Encrypted password file
        Write-Verbose -Message "Checking for required password file"
        IF (!(Test-Path -Path "$Path\O365cred.txt"))
            {
                $Pass = Read-Host -Prompt "Enter your tenant password" -AsSecureString `
                    | ConvertFrom-SecureString
            }
        else
            {
                Write-Host -ForegroundColor Green "Password File Found"
                $Pass = Get-Content "$Path\O365cred.txt" | ConvertTo-SecureString
            }
## Create Cred variable
    Write-Verbose -Message "Creating Credential varialble from files"
    $Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass

## Connect of O365 with error checks
    ##Connect MSOnline (Azure Active Directory)
        Write-Verbose -Message "Connecting MSOnline (Office 365)"
        ## Office 365
            Try 
                {
                    Write-Verbose -Message "Checking if already connected to MSOnline (Azure Active Directory)"
                    Get-MsolDomain -ErrorAction Stop > $null
                    Write-Host -ForegroundColor Green "Already Connected to MSOnline (Azure Active Directory)"
                }
            Catch
                {
                    Try 
                        {
                            Write-Verbose -Message "Not connected to MSOnline (Azure Active Directory)"
                            Write-Verbose -Message "$(Get-Date -f o) Importing Module MSOline"
                            Import-Module MSOnline -DisableNameChecking -ErrorAction Stop
                            Write-Verbose -Message "$(Get-Date -f o) Connecting to MSOL Service"
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
                            Write-Verbose -Message "Connecting MsolService"
                            Connect-MsolService -Credential $cred -ErrorAction Stop
                            Get-MsolDomain -ErrorAction Stop > $null
                            Write-Host -ForegroundColor Green "Connected to MSOnline (Azure Active Directory)"
                        }
                    Catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException] 
                        {
                            Write-Error -Message "Error. Failed to Connect to MsolService because bad username or password." -ErrorAction Stop
                        }
                    Catch 
                        {
                            #$_ | fl * -Force
                            Write-Error -Message "Error. Failed to connect to MsolService because $_" -ErrorAction Stop
                        }
                    if ($ShowProgress) 
                        {
                            Write-Verbose -Message "$(Get-Date -f o) Listing MSOL Domains"
                            Get-MsolDomain | ft -AutoSize
                        }
                }

            
    ##Connect SharePoint Online
        Write-Verbose -Message "Connecting SharePoint Online"
        ##SharePoint
            if (-not $SkipSharePoint) 
                {
                    Try 
                        {
                            Try 
                                {
                                    Write-Verbose -Message "Checking if already connected to SharePoint Online"
                                    Get-SPOsite -ErrorAction Stop > $null
                                    Write-Host -ForegroundColor Green "Already Connected to SharePoint Online"
                                }
                            Catch
                                {
                                    Write-Verbose -Message "Not connected to SharePoint Online"
                                    Write-Verbose -Message "$(Get-Date -f o) Importing SharePoint Module"
                                    Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
                                    $spAdminName = ((Get-MsolDomain | where Name -match 'onmicrosoft').Name).split('.')[0]
                                    Write-Verbose -Message "$(Get-Date -f o) This is the SP Admin Domain: '$spAdminName'"
                                }
                        }
                    Catch 
                        {
                            Write-Verbose -Message "SharePoint Online Module not found."
                            $temp = Read-Host "Launch browser to download SharePoint Online Management Shell? [Y]/N"
                            If ($temp -eq $null -OR $temp -eq "" -OR $temp -eq "Y")
                                {
                                    Start-Process https://www.microsoft.com/en-us/download/details.aspx?id=35588
                                    exit      
                                }
                            Else
                                {
                                    Write-Error -Message "Error. Failed to import the module 'Microsoft.Online.SharePoint.PowerShell' because $_" -ErrorAction Stop
                                }
                        }
                    Try 
                        {
                            Write-Verbose -Message "$(Get-Date -f o) Connecting SP Online Service"
                            Connect-SPOService -Url https://$spAdminName-admin.sharepoint.com -Credential $cred -ErrorAction Stop
                            Get-SPOSite -ErrorAction Stop > $null
                            Write-Host -ForegroundColor Green "Connected to SharePoint Online"
                        }
                    Catch 
                        {
                            Write-Error -Message "Error. Cannot connect to 'SPOService' because $_" -ErrorAction Stop
                        }
                    if ($ShowProgress) 
                        {
                            Write-Verbose -Message "$(Get-Date -f o) Listing SP Online Sites"
                            Get-SPOSite | ft -AutoSize
                        }
                } 
            else   
                {
                    Write-Verbose -Message "Skipping SharePoint"
                }
    ##Connect Skype for Business
        Write-Verbose -Message "Connecting Skype for Business"
        if (-not $SkipSkype) 
            {   
                Try 
                    {
                        Write-Verbose -Message "Checking if already connected to Skype for Business"
                        If ((Get-PSSession | Where-Object {$_.ComputerName -like "admin1a*" -AND $_.State -like "Opened"}))
                            {
                                Write-Host -ForegroundColor Green "Already Connected to Skype for Business"
                                Write-Verbose -Message "Removing broken PSSessions for Skype for Business"
                                $PSSessionIDs = Get-PSSession `
                                    | Where-Object {$_.ComputerName -like "admin1a*" -AND $_.State -notlike "Opened"} `
                                    | Select-Object Id
                                If ($PSSessionIDs -eq $null -OR $PSSessionIDs -eq "")
                                    {
                                        Write-Verbose -Message "No broken sessions to remove"
                                    }
                                else 
                                    {
                                        Foreach ($PSessionID in $PSSessionIDs)
                                            {Remove-PSSession $PSessionID}
                                        Write-Verbose -Message "All broken Skype for Business sessions removed"
                                    }
                            }
                        else 
                            {
                                Write-Verbose -Message "$(Get-Date -f o) Importing Module Skype for Business Online Connector"
                                Import-Module SkypeOnlineConnector -DisableNameChecking -ErrorAction Stop
                            }
                    }
                Catch 
                    {
                        Write-Verbose -Message "SkypeOnlineConnector Module not found."
                        $temp = Read-Host "Launch browser to download Skype for Business Online, Windows PowerShell Module? [Y]/N"
                            If ($temp -eq $null -OR $temp -eq "" -OR $temp -eq "Y")
                                {
                                    Start-Process https://www.microsoft.com/en-us/download/details.aspx?id=39366
                                    exit      
                                }
                            Else
                                {
                                    Write-Error -Message "Error. Failed to import the module 'SkypeOnlineConnector' because $_" -ErrorAction Stop
                                }
                    }
                Try 
                    {
                        Write-Verbose -Message "$(Get-Date -f o) Creating Skype for Business Session"
                        $skypeSession = New-CsOnlineSession -Credential $cred -ErrorAction Stop
                    }
                Catch 
                    {
                        Write-Warning "$(Get-Date -f o) Error. Cannot connect to Skype for Business because $_"
                    }
                Try 
                    {
                        Write-Verbose -Message "$(Get-Date -f o) Importing PSSession for Skype for Business"
                        $null = Import-PSSession $skypeSession -AllowClobber -ErrorAction Stop
                        Write-Verbose -Message "Successfully imported PSSession for Skype for Business."
                    }
                Catch 
                    {
                        Write-Error -Message "Failed to import PSSession for Skype for Business." -ErrorAction Continue
                    }
            
                if ($ShowProgress) 
                    {
                        if (Get-Command -Name Get-CsMeetingConfiguration -ErrorAction SilentlyContinue) 
                            {
                                Write-Verbose -Message "Successfully connected to Skype for Business."
                            } 
                        else 
                            {
                                Write-Warning -Message "Error. Did not connect to Skype for Business."
                            }
                    }
                    If ((Get-PSSession | Where-Object {$_.ComputerName -like "admin1a*" -AND $_.State -like "Opened"}))
                        {
                            Write-Host -ForegroundColor Green "Connected to Skype for Business"
                        }
                    else 
                        {
                            Write-Host -ForegroundColor Red "Failed to connect to Skype for Business"
                        }
            } 
        else 
            {
                Write-Verbose -Message "Skipping Skype for Business"
            }
     
    ##Connect Exchange Online
        Write-Verbose -Message "Connecting Exchange Online"
        #$Exchange
            if (-not $SkipExchange) 
                {
                    Write-Verbose -Message "Checking if already connected to Exchange Online"
                    If ((Get-PSSession | Where-Object {$_.ComputerName -like "outlook*" -AND $_.State -like "Opened"}))
                        {
                            Write-Host -ForegroundColor Green "Already Connected to Exchange Online"
                            Write-Verbose -Message "Removing broken PSSessions for Exchange"
                            $PSSessionIDs = Get-PSSession `
                                | Where-Object {$_.ComputerName -like "outlook*" -AND $_.State -notlike "Opened"} `
                                | Select-Object Id
                            If ($PSSessionIDs -eq $null -OR $PSSessionIDs -eq "")
                                {
                                    Write-Verbose -Message "No broken sessions to remove"
                                }
                            else 
                                {
                                    Foreach ($PSessionID in $PSSessionIDs)
                                        {Remove-PSSession $PSessionID}
                                    Write-Verbose -Message "All broken Exchange sessions removed"
                                }
                        }
                    else 
                        {
                            Try
                                {
                                    Write-Verbose -Message "$(Get-Date -f o) Creating Exchange Session"
                                    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange `
                                        -ConnectionUri "https://outlook.office365.com/powershell-liveid/" `
                                        -Credential $cred `
                                        -Authentication "Basic" `
                                        -AllowRedirection `
                                        -ErrorAction Stop
                                    Write-Verbose -Message "Successfully created Exchange Session."      
                                }
                            Catch 
                                {
                                    Write-Error -Message "Error. Cannot talk to Exchange because $_" -ErrorAction Stop
                                }
                            Try 
                                {
                                    Write-Verbose -Message "$(Get-Date -f o) Importing PSSession for Exchange"
                                    $null = Import-PSSession $exchangeSession -AllowClobber -ErrorAction Stop -DisableNameChecking
                                    Write-Verbose -Message "Imported PSSession for Exchange."
                                }
                            Catch 
                                {
                                    Write-Error -Message "Error. Cannot Import PSSession for Exchange because $_" -ErrorAction Stop
                                }
                            if ($ShowProgress) 
                                {
                                    Write-Verbose -Message "$(Get-Date -f o) Listing Accepted Domains"
                                    Get-AcceptedDomain | Format-Table -Property DomainName, DomainType, IsValid -AutoSize
                                }
                            If ((Get-PSSession | Where-Object {$_.ComputerName -like "outlook*" -AND $_.State -like "Opened"}))
                                {
                                    Write-Host -ForegroundColor Green "Connected to Exchange"
                                }
                            else 
                                {
                                    Write-Host -ForegroundColor Red "Failed to connect to Exchange"
                                }
                        }
                }
            else 
                {
                    Write-Verbose -Message "Skipping Exchange"
                }
    ##Connect Security & Compliance Center
        Write-Verbose -Message "Connecting Security & Compliance Center"
        #$Exchange
            if (-not $SkipSecurity) 
                {
                    If ((Get-PSSession | Where-Object {$_.ComputerName -like "*compliance.protection.outlook*" -AND $_.State -like "Opened"}))
                        {
                            Write-Host -ForegroundColor Green "Already Connected to Security & Compliance Center Online"
                            Write-Verbose -Message "Removing broken PSSessions for Security & Compliance Center"
                            $PSSessionIDs = Get-PSSession `
                                | Where-Object {$_.ComputerName -like "*compliance.protection.outlook*" -AND $_.State -notlike "Opened"} `
                                | Select-Object Id
                            If ($PSSessionIDs -eq $null -OR $PSSessionIDs -eq "")
                                {
                                    Write-Verbose -Message "No broken sessions to remove"
                                }
                            else 
                                {
                                    Foreach ($PSessionID in $PSSessionIDs)
                                        {Remove-PSSession $PSessionID}
                                    Write-Verbose -Message "All broken Exchange sessions removed"
                                }
                        }
                    else 
                        {
                             Try 
                                {
                                    Write-Verbose -Message "$(Get-Date -f o) Creating Security & Compliance Center Session"
                                    $securitySession = New-PSSession -ConfigurationName Microsoft.Exchange `
                                        -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid/" `
                                        -Credential $cred `
                                        -Authentication "Basic" `
                                        -AllowRedirection `
                                        -ErrorAction Stop
                                        Write-Verbose -Message "Successfully created Exchange Security & Compliance Session."      
                                }
                            Catch 
                                {
                                    Write-Error -Message "Error. Cannot talk to Exchange Security & Compliance because $_" -ErrorAction Stop
                                }
                            Try 
                                {
                                    Write-Verbose -Message "$(Get-Date -f o) Importing PSSession for Security & Compliance Center"
                                    $securitySession = Import-PSSession $securitySession -Prefix cc -AllowClobber -ErrorAction Stop -DisableNameChecking
                                }
                            Catch 
                                {
                                    Write-Warning "Error. Cannot Import PSSession for Security & Compliance Center because $_"
                                }                    
                            if ($ShowProgress) 
                                {
                                    if (Get-Command -Name Get-CsMeetingConfiguration -ErrorAction SilentlyContinue) 
                                        {
                                            Write-Verbose -Message "Successfully connected to Security & Compliance Center."
                                        } 
                                    else 
                                        {
                                            Write-Warning -Message "Error. Did not connect to Security & Compliance Center."
                                        }
                                }
                            If ((Get-PSSession | Where-Object {$_.ComputerName -like "*compliance.protection.outlook*" -AND $_.State -like "Opened"}))
                                {
                                    Write-Host -ForegroundColor Green "Connected to Security & Compliance Center"
                                }
                            else 
                                {
                                    Write-Host -ForegroundColor Red "Failed to connect to Security & Compliance Center"
                                }
                        }
                    
                } 
            else 
                {
                    Write-Verbose -Message "Skipping Security & Compliance Center"
                }