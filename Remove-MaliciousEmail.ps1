#####################################
## Author: James Tarran // Techary ##
#####################################

# ---------------------- ELEVATE ADMIN ----------------------

<#
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

}  #>

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
            write-host "`nExchange online Management exists"
        }
    else
        {
            write-host -ForegroundColor red "`nExchange oneline management does not exist. Installing..."

            Set-PSRepository -Name "PSgallery" -InstallationPolicy Trusted
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
            import-module ExchangeOnlineManagement

        }

    Connect-IPPSSession -InformationAction ignore -ErrorAction SilentlyContinue

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

    $Script:Subject = read-host "`nEnter the subject of the email. It is recommeneded this is pasted in for accuracy"

    get-EmailSender

}

function get-EmailSender {

    $Script:SenderAddress = read-host "`nEnter the sender of the email. It is recommeneded this is pasted in for accuracy"

    get-infoConfirmation

}

function get-infoConfirmation {

    do {
        $confirm = read-host "`nSubject entered is: $script:subject `nEmail enterted is $script:SenderAddress `nIs this correct? Y/N"
        switch ($confirm)
                {
                    Y {}
                    N {get-EmailSubject}
                    default {"You didn't enter an expected response, you idiot."}
                }
            } until ($confirm -eq 'Y')

}

function get-ContentSearchStatus {

    while ((get-complianceSearch -identity $Script:RandomIdentity).status -ne "Completed")
        {

            countdown -timeSpan 1

        }

        $script:items = (get-complianceSearch -identity $Script:RandomIdentity).items
        write-host "`n$script:items email(s) found"

        if ($script:items -ne 0)
            {

                write-host -NoNewline "`n`nDeleting $script:Items email(s), please wait"
                new-ComplianceSearchAction -searchname $Script:RandomIdentity -purge -confirm:$false | Out-Null
                remove-ContentSearchResults

            }
        else
            {

                write-host "No emails found, please confirm the sender address and subject of the email"

            }



}

function remove-ContentSearchResults {

    $delete = read-host "Do you want to delete all $script:items emails? Y/N"

    if ($delete -eq "Y")
        {

            while ((get-complianceSearchAction -identity "$Script:RandomIdentity`_purge").status -ne "Completed")
                {

                    countdown -timeSpan 1

                }

            write-host "`nDeletion of $script:items emails complete!"
        }
    elseif ($delete -eq "N")
        {

        }
    else
        {

            Write-Output "Y/N not entered. Plese try again"
            remove-ContentSearchResults

        }


}


print-TecharyLogo

connect-ComplianceCentre

Write-Output "`n`n"

Write-Warning "Ensure you have the subject of the email. `n`nThis can be obtained via the email headers."

get-EmailSubject

$Script:RandomIdentity = get-random -Maximum 999999

new-complianceSearch -name $Script:RandomIdentity -ExchangeLocation all -ContentMatchQuery "(From:$script:SenderAddress) AND (Subject:$script:subject)" | Start-ComplianceSearch

write-host -NoNewline "`nSearching, please wait"

get-ContentSearchStatus

Disconnect-ExchangeOnline -Confirm:$false -InformationAction ignore -ErrorAction SilentlyContinue