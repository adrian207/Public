<#
.SYNOPSIS
    Check allowed senders to a mail enabled group
.DESCRIPTION
    Check allowed senders to a mail enabled group
.PARAMETER Identity
    Targeted mail enabled Group Name
.EXAMPLE
    Pull list of allowed senders for target group
    
    Get-ADGroupAllowedSenders.ps1 -identity groupname
.NOTES
    Created by Chris Lee
    Date August 3, 2017
.LINK
    GitHub: 
    Blogger:    
#>

[Cmdletbinding()]
Param
(
    [parameter(Mandatory=$TRUE,Position=1)]
    [string]
    $Identity
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

## Load list of allowed senders
    $Senders = Get-ADGroup -Identity $Identity -Properties authOrig `
        | Select-Object -ExpandProperty authOrig
## Convert DN List to display name    
    ForEach ($Sender in $Senders)
        {
            Get-ADNameTranslation.ps1 -InputType DN -Name $Sender -OutputType display
        }