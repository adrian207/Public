<#
.SYNOPSIS
    Run Azure (DirSync) on dedicated remote server
.DESCRIPTION
    Run Azure (DirSync) on dedicated remote server
.PARAMETER Type
    Type of Sync to complete

    o Delta - Changes only
    o Full - Complete re-sync
.PARAMETER ADSyncPath
    Enter path for ADSync modeule if different then default
.PARAMETER Server
    Target server that has Azure Sync installed
.EXAMPLE
    Execute a full Azure Sync
    
    Invoke-O365AzureSync.ps1 -Server server -Type Full
.EXAMPLE
    Execute a change only Azure Sync
    
    Invoke-O365AzureSync.ps1 -Server server -Type Delta
.NOTES
    Created by Chris Lee
    Date September 6, 2016
.LINK   
#>

[Cmdletbinding()]
Param
(
    [string]
    [Parameter(Mandatory=$true)]
    [ValidateSet('Delta','Full')]
    $type,
    [string]
    $ADSyncPath= 'C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync\ADSync.psd1',
    [string]
    $Server
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

Write-Verbose -Message "Opening PSSession to $Server"
$Session = New-PSSession -ComputerName $Server
Write-Verbose -Message "Checking for ADSync.psd1 at $ADSyncPath"
If (Invoke-command -Session $Session -ArgumentList $ADSyncPath -scriptblock { param ($ADSyncPath) Test-Path $ADSyncPath})
    {
        Write-Verbose -Message "ADSync Module found"
        Invoke-Command -Session $Session -ArgumentList $ADSyncPath -scriptblock { param ($ADSyncPath) Import-Module $ADSyncPath}
        Write-Verbose -Message "ADSync Module Loaded"
        If ($type -eq "Delta")
            {
                Write-Verbose -Message "Executing Change Only (Delta) Sync"
                Invoke-Command -Session $Session {Start-ADSyncSyncCycle -PolicyType Delta}
            }
        Else
            {
                Write-Verbose -Message "Executing Full Sync"
                Invoke-Command -Session $Session {Start-ADSyncSyncCycle -PolicyType Initial}
            }
    }
Else
    {
        Verbose-Error "Unable to locate ADSync.psd1 at $ADSyncPath"
    }
Write-Verbose -Message "Removing PSSession to $Server"
Remove-PSSession $Session