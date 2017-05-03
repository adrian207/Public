<#
.SYNOPSIS
    Remove open O365 PSSessions
.DESCRIPTION
    Remove open O365 PSSessions
.EXAMPLE
    Remove-O365PSSession.ps1
.NOTES
    Created by Chris Lee
    Date April 20, 2017
.LINK 
    GitHub: https://github.com/clee1107/Public/blob/master/O365/Remove-O365PSSession.ps1
    Blogger:   
#>

#################################
## DO NOT EDIT BELOW THIS LINE ##
#################################

##Close open All O365 product Sessions
Write-Verbose -Message "Removing all PSSessions"
Get-PSSession | Remove-PSSession
Write-Verbose -Message "Disconnecting SharePoint Online"
$ExitPowershell = Read-Host -Prompt "Disconnect from O365 (Will close current Powershell Window) [Y]/N"
If ($ExitPowershell -eq "Y" -OR $ExitPowershell -eq $null -OR $ExitPowershell -eq "")
    {
        stop-process -Id $PID
    }