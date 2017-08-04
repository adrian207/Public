<#
.SYNOPSIS
    Set allowed senders to a mail enabled group
.DESCRIPTION
    Set allowed senders to a mail enabled group
.PARAMETER Identity
    Targeted mail enabled Group Name
.Parameter NewSender
    Enter user email (UPN) to add to allowed senders
.Parameter Server
    Enter target server, if not provided targets domain
.EXAMPLE
    Add sender to target mail enabled group
    
    Get-ADGroupAllowedSenders.ps1 -identity groupname -sender senderusername
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
    $Identity,
    [parameter(Mandatory=$TRUE,Position=1)]
    [string]
    $NewSender,
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

## Convert Sender to DN
    Write-Verbose -Message "Convert provided user's UPN to Distinguished Name (DN) and Display Name format"
    $NewSenderDN = Get-ADNameTranslation.ps1 -InputType UPN -Name $NewSender -OutputType DN
    $NewSenderDisplay = Get-ADNameTranslation.ps1 -InputType DN -Name $NewSenderDN -OutputType display

## Load list of allowed senders
    Write-Verbose -Message "Pull current members of $Identity"
    $Senders = Get-ADGroup -Identity $Identity -Properties authOrig -Server $Server `
        | Select-Object -ExpandProperty authOrig
## Check if user is already member of allowed senders    
    Write-Verbose -Message "Validate $NewSenderDisplay not a member of allowed senders for $Identity"
    ForEach ($Sender in $Senders)
        {
            If ($Sender -eq $NewSenderDN)
                {
                    Write-Host -ForegroundColor Yellow "$NewSenderDisplay already member of $Identity"
                    exit
                }
        }
## Add NewSender as a member
    Write-Verbose -Message "Adding $NewSenderDisplay to $Identity"
    Write-Host -ForegroundColor Green "Adding $NewSenderDisplay to $Identity"
    Set-ADGroup -Identity $Identity -add @{authOrig=$NewSenderDN}
    