<#
Version: .01
Date Created: August 10, 2016
Created by: Chris Lee
Purpose:
    Create Aliases for CAPS students
Resources:
    https://github.com/jay0lee/GAM/wiki
Revision Notes:
    .01 - Intial Setup
#>

#region Menu
function Show-Menu
{
     param (
           [string]$Title = 'GAFE Alias Management'
     )
     cls
     Write-Host "===================== $Title ====================="
     #Write-Host "12345678901234567890123456789012345678901234567890"
     
     Write-Host ""
     Write-Host "Individual"
     Write-Host "----------------------------"
     #Write-Host "12345678901234567890123456789012345678901234567890"
     Write-Host "      1:        Add Alias"
     Write-Host "      2:        Remove Alias"
     Write-Host "      3:        Report Alias"
     Write-Host ""
     Write-Host "Mass via CSV"
     Write-Host "--------------------------------------------"
     #Write-Host "12345678901234567890123456789012345678901234567890"
     Write-Host "      4:         Add Alias"
     Write-Host "      5:         Remove Alias"
     Write-Host "      6:         Report Alias"
     Write-Host "      7:         CSV Template"
     Write-Host ""
     Write-Host ""
     Write-Host "R:  Press 'R' for Read-Me"
     Write-Host "Q:  Press 'Q' to quit."
     Write-Host ""
     Write-Host ""
}
#endregion

#region Execute Menu Selectio
do
{
    Show-Menu
    $input = Read-Host "Please make a selection"   
    switch ($input)
        {
           #region Individual
            '1' { #Set Alias
                    $samid = Read-Host -Prompt "Input username:"
                    $alias = Read-Host -Prompt "Input desired alias:"
                    $Session = New-PSSession -ComputerName SVC01
                    Invoke-Command -Session $Session -ArgumentList $alias, $samid -scriptblock { 
                        param ($alias, $samid) C:\GAM\GAM.EXE create alias $using:alias user $using:samid 
                        }
                    Remove-PSSession $Session
                }
            '2' { #Remove Alias
                    $alias = Read-Host -Prompt "Input alias to remove:"
                    $Session = New-PSSession -ComputerName SVC01
                    Invoke-Command -Session $Session -ArgumentList $alias, $samid -scriptblock { 
                        param ($alias, $samid) C:\GAM\GAM.EXE delete alias $using:alias
                        }
                    Remove-PSSession $Session
                }
            '3' { #Report Alias
                    $alias = Read-Host -Prompt "Input alias to remove:"
                    $Session = New-PSSession -ComputerName SVC01
                    Invoke-Command -Session $Session -ArgumentList $alias, $samid -scriptblock { 
                        param ($alias, $samid) C:\GAM\GAM.EXE info alias $using:alias 
                        }
                    Remove-PSSession $Session
                }
           #endregion
           #region Mass
            '4' { 
                    $File = Read-Host -Prompt "Enter file with full path name (Default: C:\Users\[user]\desktop\GAFEAlias.csv)"
                        #Checks if user hit Enter for Default 
                        if($File -eq $null -Or $File -eq ""){$File = $Desktop+"\GAFEAlias.csv"}
                    $Users = Import-Csv $File -Delimiter ","
                    $Session = New-PSSession -ComputerName SVC01
                    ForEach ($User in $Users) {
                        $alias = $User.Alias
                        $samid = $User.name
                        Invoke-Command -Session $Session -ArgumentList $alias, $samid -scriptblock { 
                            param ($alias, $samid) C:\GAM\GAM.EXE create alias $using:alias user $using:samid 
                            }
                        }
                    Remove-PSSession $Session
                }
            '5' { 
                    $File = Read-Host -Prompt "Enter file with full path name (Default: C:\Users\[user]\desktop\GAFEAlias.csv)"
                        #Checks if user hit Enter for Default 
                        if($File -eq $null -Or $File -eq ""){$File = $Desktop+"\GAFEAlias.csv"}
                    $Users = Import-Csv $File -Delimiter ","
                    $Session = New-PSSession -ComputerName SVC01
                    ForEach ($User in $Users) {
                        $alias = $User.Alias
                        Invoke-Command -Session $Session -ArgumentList $alias -scriptblock { 
                            param ($alias) C:\GAM\GAM.EXE delete alias $using:alias
                            }
                        }
                    Remove-PSSession $Session
                }
            '6' { 
                    $File = Read-Host -Prompt "Enter file with full path name (Default: C:\Users\[user]\desktop\GAFEAlias.csv)"
                        #Checks if user hit Enter for Default 
                        if($File -eq $null -Or $File -eq ""){$File = $Desktop+"\GAFEAlias.csv"}
                    $Users = Import-Csv $File -Delimiter ","
                    $Session = New-PSSession -ComputerName SVC01
                    ForEach ($User in $Users) {
                        $alias = $User.Alias
                        Invoke-Command -Session $Session -ArgumentList $alias -scriptblock { 
                            param ($alias) C:\GAM\GAM.EXE info alias $using:alias
                            }
                        }
                    Remove-PSSession $Session
                }
            '7' {
                    "User,Alias" | select 'Name', 'Alias' | Export-Csv $Desktop\GAFEAlias.csv -NoTypeInformation
                }
           #endregion
           #endregion 
           'R' {
                    cls
                    'You chose option #R'
               }
        }
    pause   
}
until ($input -eq 'q')
#endregion