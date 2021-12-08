#####################################
## Author: James Tarran // Techary ##
#####################################

# ---------------------- ELEVATE ADMIN ---------------------- 

param([switch]$Elevated)

function Test-Admin {
  $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
  $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

if ((Test-Admin) -eq $false)  {
    if ($elevated) 
    {
        # tried to elevate, did not work, aborting
    } 
    else {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
}

exit

}

# Prints 'Techary' in ASCII
function print-TecharyLogo {
        
    $logo = "
     _______        _                      
    |__   __|      | |                     
       | | ___  ___| |__   __ _ _ __ _   _ 
       | |/ _ \/ __| '_ \ / _`` | '__| | | |
       | |  __/ (__| | | | (_| | |  | |_| |
       |_|\___|\___|_| |_|\__,_|_|   \__, |
                                      __/ |
                                     |___/ 
"

write-host -ForegroundColor Green $logo

}

function connect-ComplianceCentre {

    if (Get-Module -ListAvailable -Name ExchangeOnlineManagement) 
        {
            write-host " "
            write-host "Exchange online Management exists"
        }
    else 
        {

            Set-PSRepository -Name "PSgallery" -InstallationPolicy Trusted
            Install-Module -Name ExchangeOnlineManagement
            import-module ExchangeOnlineManagement

        }

    Connect-IPPSSession

}

function CountDown() {
    param($timeSpan)

    while ($timeSpan -gt 0)
  {
    Write-Host '.' -NoNewline
    $timeSpan = $timeSpan - 1
    Start-Sleep -Seconds 1
  }
}

function get-EmailSubject {
            
    write-host " "

    $Script:Subject = read-host "Enter the subject of the email. It is recommeneded this is pasted in for accuracy"

    do {
    $confirm = read-host "The subject entered is: `n$script:subject `nIs this correct? Y/N"
    switch ($confirm)
            {
                Y {get-EmailSender}
                N {get-EmailSubject}
                default {"You didn't enter an expected response, you idiot."}
            } 
        } until ($confirm -eq 'Y')
}

function get-EmailSender {

    write-host " "

    $Script:SenderAddress = read-host "Enter the sender of the email. It is recommeneded this is pasted in for accuracy"

    do {
    $confirm = read-host "The address entered is: `n$script:SenderAddress `nIs this correct? Y/N"
    switch ($confirm)
            {
                Y {}
                N {get-EmailSender}
                default {"You didn't enter an expected response, you idiot."}
            } 
        } until ($confirm -eq 'Y')

}

function get-ContentSearchStatus {

    countdown -timeSpan 10

    if ((get-complianceSearch).status -ne "Completed")
        {

            get-ContentSearchStatus

        }
    else 
        {

            write-host "`nDeleting, please wait"
            remove-ContentSearchResults

        }

}

function remove-ContentSearchResults {

    countdown -timeSpan 10

    if ((get-complianceSearchAction).status -ne "Completed")
    {

        remove-ContentSearchResults

    }
else 
    {

        write-host "`nDeletion complete!"

    }

    
}

print-TecharyLogo

connect-ComplianceCentre

Write-Warning "Ensure you have the following information: `n`nThe subject of the email. `nThe original sender of the email. `n`nThese can be attained via the email headers."

get-EmailSubject

$Script:RandomIdentity = get-random -Maximum 999999

$ComplianceSearch = new-complianceSearch -name $Script:RandomIdentity -ExchangeLocation all -ContentMatchQuery "(From:$script:SenderAddress) AND (Subject:'$script:subject')"

Start-ComplianceSearch -Identity $ComplianceSearch.Identity

write-host "`nSearching, please wait"

get-ContentSearchStatus

disconnect-ExchangeOnline -Confirm:$false