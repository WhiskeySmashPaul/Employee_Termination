#region Modules Required

Add-Type -AssemblyName System.Web

#endregion Modules Required

#region Adjustable Variables

$TerminationOU = "<Enter OU where to store terminated employees>"

#endregion Adjustable Variables

#region Static Variables

#Prompt to enter users username for termination workflow
$FirstChoice = '&Yes', '&No', '&Quit'
$choices  = '&Yes', '&No'
$adName = Read-Host "Please enter username to terminate"
$QuestionUserName = "Is users username $adName"
$TitleUserName = "Username Confirmation"
$HomeFolderPath = "<Enter your Homefolder path here>"
$ArchiveFolderPath = "<Enter your HomeFolder Archive Path here>"

#endregion Static Variables

#region Create Function

function Terminate-Employee {

#region Verify Employee Terminating

    #Verifies users first name with Administartator, yes continues no reprompts for first name until recieves yes#
    do{
        $decision2 = $Host.UI.PromptForChoice($TitleUserName, $QuestionUserName, $Firstchoice, 1)
            if ($decision2 -eq 0) {
                Write-Host 'Confirmed' -ForegroundColor Green
            }
            elseif ($decision2 -eq 2){
                Write-Host 'Quiting the program now.' -ForegroundColor Yellow
                Pause 5
                Exit
            }
            else {
                Write-Host 'Input correct username' -ForegroundColor Red
                $adName = Read-Host -Prompt "Input users correct username"
            }
    } until ($decision2 -eq 0)

#endregion Verify Employee Terminating 

#region Generate Secure Password

    #creates a random secure passowrd#
    $password = [System.Web.Security.Membership]::GeneratePassword((Get-Random -Minimum 18 -Maximum 26), 3)
    $secPw = ConvertTo-SecureString -String $password -AsPlainText -Force
    Write-Host "Changing user password" -ForegroundColor DarkCyan
    Set-ADAccountPassword -Identity $adName -NewPassword $secPw -Reset

#endregion Generate Secure Password

#region Remove All Memberships

    #Find all group memberships and remove them except TermedUsersGroup and move to TerminatedUsers OU in AD
    Write-Host "Removing memberships from user" -ForegroundColor DarkCyan
    $Groups = Get-ADPrincipalGroupMembership -Identity $adName | Where-Object {$_.Name -notlike "Domain Users"}

    foreach($Group in $Groups)
    {
        Remove-ADPrincipalGroupMembership -identity $adName -memberof $Group -Confirm:$false
    }

#endregion Remove All Memberships

#region Move User OU

    #Moves user to the terminations OU
    Write-Host "Moving user to the Terminations OU" -ForegroundColor DarkCyan
    Get-ADUser $adName | Move-ADObject -TargetPath $TerminationOU -Confirm:$false

#endregion Move User OU

#region Add Termination Date to Description

    #Adds users termination date to description in AD
    Write-Host "Updating AD User description fields" -ForegroundColor DarkCyan
    $userDescription = get-aduser $adName -Properties description | Select-Object -ExpandProperty Description
    $terminationDate = Get-Date -Format "MM-dd-yyyy"
    Set-ADUser -Identity $adName -Description ("$userDescription - Terminated $terminationDate")

#endregion Add Termination Date to Description

#region Disable Account

    #Disables AD account
    Write-Host "Disabling AD Account" -ForegroundColor DarkCyan
    Disable-ADAccount -Identity $adName

#endregion Disable Account

#region Move Home Folder
##NOTE-Moves the employees Home Folder to a sub folder in the directory for access for 30 days before being deleted by another automation script##

    #Renames L drive to show terminated date and moves to terminated employee folder
    Write-Host "Moving user personal folder to the ArchivedEmployee folder" -ForegroundColor DarkCyan
    Move-Item -Path "$HomeFolderPath\$adName" -Destination "$ArchiveFolderPath\$adName - Terminated $terminationDate"
    
#endregion Move Home Folder

#region Loop for another user
    
    #region Loop Variables
    ##Variables for looping the script to terminate another user##
    $QuestionReprompt = "Would you like to terminate another user?"
    $TitleRePrompt = "Terminate another user"
    $decision1 = $Host.UI.PromptForChoice($TitleRePrompt, $QuestionReprompt, $choices, 1)

    #endregion Loop Variables

    if ($decision1 -eq 0) {
       Write-Host 'Proceeding to terminate another user' -ForegroundColor Green
       Terminate-Employee
    } else {
        Write-Host 'Finished terminating user(s)' -ForegroundColor Green
        Start-Sleep -Seconds 3
        Exit    
    }

#endregion Loop for another user

}

#endregion Create Function

#region Call Function

Terminate-Employee

#endregion Call Function
