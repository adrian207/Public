<#
.SYNOPSIS
    Remove user from allowed senders of a mail enabled group
.DESCRIPTION
    Remove user from allowed senders of a mail enabled group
.PARAMETER Identity
    Targeted mail enabled Group Name
.Parameter RemoveSender
    Enter user email (UPN) to remove from allowed senders
.Parameter Server
    Enter target server, if not provided targets domain
.EXAMPLE
    Remove sender to target mail enabled group
    
    Remove-ADGRoupAllowedSenders.ps1 -identity groupname -sender senderusername
.NOTES
    Created by Chris Lee
    Date August 3, 2017

    Requires Get-ADNameTranslation.ps1
.LINK
    GitHub: 
    Blogger:    
#>

[Cmdletbinding()]
Param
(
    [parameter(Mandatory=$TRUE,Position=1)]
    [string]
    $Identity,
    [parameter(Mandatory=$TRUE,Position=1)]
    [string]
    $RemoveSender,
    [String]
    $Server = (Get-ADDomain | Select-Object -ExpandProperty DNSRoot)
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

## Check if Get-ADNameTranslation.ps1 present in same directory
    # Pull Exection directory
        $MyRootPath = Get-Item -Path $MyInvocation.MyCommand.Path
    # Convert to usable distory value
        $MyRootPath = "$($MyRootPath.DirectoryName)"
    # Check for Get-ADNameTranslation.ps1 present in MyRootPath directory
        Write-Verbose -Message "Checking for required sub-script Get-ADNameTranslation.ps1"
        If (!(Test-Path $MyRootPath\Get-ADNameTranslation.ps1))
            {
                Throw "Get-ADNameTranslation.ps1 not present in $MyRootPath"
            }

## Convert Sender to DN
    Write-Verbose -Message "Convert provided user's UPN to Distinguished Name (DN) and Display Name format"
    $RemoveSenderDN = Get-ADNameTranslation.ps1 -InputType UPN -Name $RemoveSender -OutputType DN
    $RemoveSenderDisplay = Get-ADNameTranslation.ps1 -InputType DN -Name $RemoveSenderDN -OutputType display

## Load list of allowed senders
    Write-Verbose -Message "Pull current members of $Identity"
    $Senders = Get-ADGroup -Identity $Identity -Properties authOrig -Server $Server `
        | Select-Object -ExpandProperty authOrig
## Check if user member of allowed senders and remove
    Write-Verbose -Message "Checking if $RemoveSenderDisplay member of allowed senders for $Identity"
    ForEach ($Sender in $Senders)
        {
            If ($Sender -eq $RemoveSenderDN)
                {
                    Write-Verbose -Message "Removing $RemoveSenderDisplay from $Identity"
                    Write-Host -ForegroundColor Green "Removing $RemoveSenderDisplay from $Identity"
                    Set-ADGroup -Identity $Identity -remove @{authOrig=$RemoveSenderDN} -Server $Server
                    exit
                }
        }
## RemoveSender not a member
    Write-Host -ForegroundColor Yellow "$RemoveSenderDisplay not a member of $Identity"
