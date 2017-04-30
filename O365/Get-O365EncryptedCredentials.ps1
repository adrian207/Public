<#
.SYNOPSIS
    Check for O365 user and credential files, Test Connect to O365
.DESCRIPTION
    Check for O365 user and credential files:
        o user.txt - Contains O365 UPN (Optional)
        o cred.txt - Contains encrypted O365 password (Required)

    Execute PoSH connect to validate credentials
    Installs missing modules for:  Active Directory Online, Lync Online, SharePointOnline, ExchangeOnline and Office 365 for an organization.
.PARAMETER SkipSharePoint
    If no DNS for SharePoint, add this to skip SharePoint options
.PARAMETER SkipLync
    If no DNS for Lync, add this to skip Lync options
.PARAMETER SkipExchange
    If no DNS for Exchange, add this to skip Exchange options
.PARAMETER SkipSecurity
    If no DNS for Security & Compliance Center, add this to skip Security & Compliance Center options
.PARAMETER Path
    Enter alternate path to save files to, defualt is users local app data
.EXAMPLE
.NOTES
    Created by Chris Lee
    Date April 20, 2017
.LINK   
#>

[Cmdletbinding()]
Param
(
    [switch]
    $SkipSharePoint,
    [switch]
    $SkipLync,
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

##Load User name
    IF (!(Test-Path -Path "$Path\O365user.txt"))
        {
            Write-Host -ForegroundColor Red "No user file will be prompted for username for all scripts"
            Write-Host -ForegroundColor Red "Consider running Set-O365EncryptedCredentials again and provide O365 UPN"
        }
    else
        {
            Write-Host -ForegroundColor Green "User File found"
            $AdminName = Get-Content "$Path\O365user.txt"
        }

##Load Encrypted password
    IF (!(Test-Path -Path "$Path\O365cred.txt"))
        {
            Write-Host -ForegroundColor Red "No password file found"
            Write-Host -ForegroundColor Red "Run Set-O365EncryptedCredentials"
        }
    else
        {
            Write-Host -ForegroundColor Green "Password File Found"
            $Pass = Get-Content "$Path\O365cred.txt" | ConvertTo-SecureString
        }

## Create Cred variable
    $Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $AdminName, $Pass

## Connect of O365 with error checks
    ##Check for open O365 session
        Write-Verbose -Message "Checking if already connected to Office 365"
        Try
            {
                Get-MsolDomain -ErrorAction Stop > $null
                Write-Host -ForegroundColor Green "Already Connected to Office 365"
            }
        catch
            {
                ##Importing MSOnline Module
                Write-Verbose -Message "Connecting MSOnline (Office 365)"
                ## Office 365
                    Try 
                        {
                            Write-Verbose -Message "$(Get-Date -f o) Importing Module MSOline"
                            Import-Module MSOnline -DisableNameChecking -ErrorAction Stop
                            Write-Verbose -Message "$(Get-Date -f o) Connecting to MSOL Service"
                        }
                    Catch 
                        {
                            Write-Error -Message "Error. Cannot import module 'MSOnline' because $_" -ErrorAction Stop
                        }
                    Try 
                        {
                            Connect-MsolService -Credential $cred -ErrorAction Stop
                            Write-Verbose -Message "Connected MsolService"
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
                 Try
                    {
                        Get-MsolDomain -ErrorAction Stop > $null
                        Write-Host -ForegroundColor Green "Connected to Office 365"
                    }
                catch
                    {
                        Write-Error "Failed to connect to Office 365"
                        exit
                    }
            }
            
    ##Check for open Sharepoint session
        Write-Verbose -Message "Checking if already connected to SharePoint Online"
        Try 
            {
                Get-SPOsite -ErrorAction Stop > $null
                Write-Host -ForegroundColor Green "Already Connected to SharePoint Online"
            }
        Catch 
            {                  
                ##Importing SharePoint Module
                Write-Verbose -Message "Importing SharePoint Module"
                ##SharePoint
                    if (-not $SkipSharePoint) {
                        Try {
                            Write-Verbose -Message "$(Get-Date -f o) Importing SharePoint Module"
                            Import-Module Microsoft.Online.SharePoint.PowerShell -DisableNameChecking -ErrorAction Stop
                            $spAdminName = ((Get-MsolDomain | where Name -match 'onmicrosoft').Name).split('.')[0]
                            Write-Verbose -Message "$(Get-Date -f o) This is the SP Admin Domain: '$spAdminName'"
                        }
                        Catch {
                            Write-Error -Message "Error. Failed to import the module 'Microsoft.Online.SharePoint.PowerShell' because $_" -ErrorAction Stop
                        }
                        Try {
                            Write-Verbose -Message "$(Get-Date -f o) Connecting SP Online Service"
                            Connect-SPOService -Url https://$spAdminName-admin.sharepoint.com -Credential $cred -ErrorAction Stop
                            Write-Verbose -Message "Connected to 'SPOService'"
                        }
                        Catch {
                            Write-Error -Message "Error. Cannot connect to 'SPOService' because $_" -ErrorAction Stop
                        }
                        if ($ShowProgress) {
                            Write-Verbose -Message "$(Get-Date -f o) Listing SP Online Sites"
                            Get-SPOSite | ft -AutoSize
                        }
                    } else {
                        Write-Verbose -Message "Skipping SharePoint"
                    }
                Try
                    {
                        Get-SPOSite -ErrorAction Stop > $null
                        Write-Host -ForegroundColor Green "Connected to SharePoint Online"
                    }
                catch
                    {
                        Write-Host -ForegroundColor Red "Failed to connect to SharePoint Online"
                    }
            }
    ##Check for open Skype for Business session
        Write-Verbose -Message "Checking if already connected to SharePoint Online"
        If ((Get-PSSession | Where-Object {$_.ComputerName -like "admin1a*" -AND $_.State -like "Opened"}))
            {
                Write-Host -ForegroundColor Green "Already Connected to SharePoint Online"
            }
        else 
            {
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
                ##Importing Skype for Business Module
                Write-Verbose -Message "Importing Skype for Business Module"
                    if (-not $SkipSkype) {
                        Try {
                            Write-Verbose -Message "$(Get-Date -f o) Importing Module Skype for Business Online Connector"
                            Import-Module SkypeOnlineConnector -DisableNameChecking -ErrorAction Stop
                        }
                        Catch {
                            Write-Error -Message "Error. Failed to import module 'SkypeOnlineConnector' because $_" -ErrorAction Stop
                        }
                        Try {
                            Write-Verbose -Message "$(Get-Date -f o) Creating Skype for Business Session"
                            $skypeSession = New-CsOnlineSession -Credential $cred -ErrorAction Stop
                        }
                        Catch {
                            Write-Warning "$(Get-Date -f o) Error. Cannot connect to Skype for Business because $_"
                        }
                        Try {
                            Write-Verbose -Message "$(Get-Date -f o) Importing PSSession for Skype for Business"
                            $null = Import-PSSession $skypeSession -AllowClobber -ErrorAction Stop
                            Write-Verbose -Message "Successfully imported PSSession for Skype for Business."
                        }
                        Catch {
                            Write-Error -Message "Failed to import PSSession for Skype for Business." -ErrorAction Continue
                        }
                    
                        if ($ShowProgress) {
                            if (Get-Command -Name Get-CsMeetingConfiguration -ErrorAction SilentlyContinue) {
                                Write-Verbose -Message "Successfully connected to Skype for Business."
                            } else {
                                Write-Warning -Message "Error. Did not connect to Skype for Business."
                            }
                        }
                    } else {
                        Write-Verbose -Message "Skipping Skype for Business"
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
    ##Check for open Exchange session
        Write-Verbose -Message "Checking if already connected to Exchange Online"
        If ((Get-PSSession | Where-Object {$_.ComputerName -like "outlook*" -AND $_.State -like "Opened"}))
            {
                Write-Host -ForegroundColor Green "Already Connected to Exchange Online"
            }
        else 
            {
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
                ##Importing Exchange Module
                Write-Verbose -Message "Importing Exchange Module"
                #$Exchange
                    if (-not $SkipExchange) {
                        Try {
                            Write-Verbose -Message "$(Get-Date -f o) Creating Exchange Session"
                            $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange `
                                -ConnectionUri "https://outlook.office365.com/powershell-liveid/" `
                                -Credential $cred `
                                -Authentication "Basic" `
                                -AllowRedirection `
                                -ErrorAction Stop
                            Write-Verbose -Message "Successfully created Exchange Session."
                        }
                        Catch {
                            Write-Error -Message "Error. Cannot talk to Exchange because $_" -ErrorAction Stop
                        }
                        Try {
                            Write-Verbose -Message "$(Get-Date -f o) Importing PSSession for Exchange"
                            $null = Import-PSSession $exchangeSession -AllowClobber -ErrorAction Stop -DisableNameChecking
                            Write-Verbose -Message "Imported PSSession for Exchange."
                        }
                        Catch {
                            Write-Error -Message "Error. Cannot Import PSSession for Exchange because $_" -ErrorAction Stop
                        }
                        if ($ShowProgress) {
                            Write-Verbose -Message "$(Get-Date -f o) Listing Accepted Domains"
                            Get-AcceptedDomain | Format-Table -Property DomainName, DomainType, IsValid -AutoSize
                        }
                    } else {
                        Write-Verbose -Message "Skipping Exchange"
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
    ##Check for open Security & Compliance Center session
        Write-Verbose -Message "Checking if already connected to Security & Compliance Center Online"
        If ((Get-PSSession | Where-Object {$_.ComputerName -like "*compliance.protection.outlook*" -AND $_.State -like "Opened"}))
            {
                Write-Host -ForegroundColor Green "Already Connected to Security & Compliance Center Online"
            }
        else 
            {
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
                ##Importing Security & Compliance Center Module
                Write-Verbose -Message "Importing Security & Compliance Center Module"
                #$Exchange
                    if (-not $SkipSecurity) {
                        Try {
                            Write-Verbose -Message "$(Get-Date -f o) Creating Security & Compliance Center Session"
                            $securitySession = New-PSSession -ConfigurationName Microsoft.Exchange `
                                -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid/" `
                                -Credential $cred `
                                -Authentication "Basic" `
                                -AllowRedirection `
                                -ErrorAction Stop
                        }
                        Catch {
                            Write-Error -Message "Error. Failed to import module 'SkypeOnlineConnector' because $_" -ErrorAction Stop
                        }
                        Try {
                            Write-Verbose -Message "$(Get-Date -f o) Importing PSSession for Security & Compliance Center"
                            $securitySession = Import-PSSession $securitySession -Prefix cc -AllowClobber -ErrorAction Stop -DisableNameChecking
                        }
                        Catch {
                            Write-Warning "Error. Cannot Import PSSession for Security & Compliance Center because $_"
                        }                    
                        if ($ShowProgress) {
                            if (Get-Command -Name Get-CsMeetingConfiguration -ErrorAction SilentlyContinue) {
                                Write-Verbose -Message "Successfully connected to Security & Compliance Center."
                            } else {
                                Write-Warning -Message "Error. Did not connect to Security & Compliance Center."
                            }
                        }
                    } else {
                        Write-Verbose -Message "Skipping Security & Compliance Center"
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

##Close open O365 Sessions
    Write-Verbose -Message "Removing all PSSessions"
    Get-PSSession | Remove-PSSession
    Write-Verbose -Message "Disconnecting SharePoint Online"
    $ExitPowershell = Read-Host -Prompt "Disconnect from O365 (Will close current Powershell Window) [Y]/N"
    If ($ExitPowershell -eq "Y" -OR $ExitPowershell -eq $null -OR $ExitPowershell -eq "")
        {
            stop-process -Id $PID
        }