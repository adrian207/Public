<#
.SYNOPSIS
    Check allowed senders to a mail enabled group
.DESCRIPTION
    Check allowed senders to a mail enabled group
.PARAMETER Identity
    Targeted mail enabled Group Name
.Parameter Server
    Enter target server, if not provided targets domain
.EXAMPLE
    Pull list of allowed senders for target group
    
    Get-ADGroupAllowedSenders.ps1 -identity groupname
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
    
## Load list of allowed senders
    $Senders = Get-ADGroup -Identity $Identity -Properties authOrig -Server $Server `
        | Select-Object -ExpandProperty authOrig
## Convert DN List to display name    
    ForEach ($Sender in $Senders)
        {
            Get-ADNameTranslation.ps1 -InputType DN -Name $Sender -OutputType display 
        }