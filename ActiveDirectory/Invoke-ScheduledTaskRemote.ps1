<#
.SYNOPSIS
    Execute scheduled task on remote server
.DESCRIPTION
    Execute scheduled task on remote server
.PARAMETER TaskName
    Task Name on target server

    Set Param hard code if desire to execute script for single task
    $TaskName = "TaskName"
.PARAMETER TaskPath
    Path to desired task name (defualt is "\"")
.PARAMETER Server
    Target server that scheduled task is setup on
.EXAMPLE
    Execute a scheduled task on server
    
    Invoke-ScheduledTaskRemote.ps1 -TaskName task -Server server
.NOTES
    Created by Chris Lee
    Date September 6, 2016
.LINK   
    GitHub: 
    Blogger: 
#>

[Cmdletbinding()]
Param
(
    [string]
    $TaskPath = "\",
    [string]
    $TaskName,
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

## Open session to remote server
    Write-Verbose -Message "Opening PSSession to $Server"
    $Session = New-PSSession -ComputerName $Server
## Execute target task
    Write-Verbose "Executing  ScheduledTask $TaskPath$TaskName on $Server"
    Invoke-Command -Session $Session -ArgumentList $TaskPath, $TaskName -scriptblock { param ($TaskPath, $TaskName) Start-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath}
    Invoke-Command -Session $Session -ArgumentList $TaskPath, $TaskName -scriptblock { param ($TaskPath, $TaskName) Get-ScheduledTaskInfo -TaskName $TaskName -TaskPath  $TaskPath}
## Close session to remote server
    Write-Verbose -Message "Removing PSSession to $Server"
    Remove-PSSession $Session