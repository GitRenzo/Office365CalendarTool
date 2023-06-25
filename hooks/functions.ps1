# To get the path where the ps1 files we want to get are stored
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptSubDirectoryFunctions = $scriptDirectory.ToString()
$bannerFile = Join-Path -Path $scriptSubDirectoryFunctions -ChildPath "banner.ps1"


# Dot source the file to import the functions
. $bannerFile

function listUsers(){
    showBanner
    Get-MailboxFolderPermission -Identity $calendarEmailAddress
}
function addOrRemoveFromCalendar($userPrincipalName) {

    if($calendarAction -eq "add"){
        $accessRightsValue = Read-Host "Enter the access rights (Owner\Editor\Reviewer\etc): "
        Add-MailboxFolderPermission -Identity $calendarEmailAddress -User $userPrincipalName -AccessRights $accessRightsValue
            }
           
    elseif($calendarAction -eq  "remove"){
        Remove-MailboxFolderPermission -Identity $calendarEmailAddress -User $userPrincipalName
            }
}

function addOrRemoveUsersFromCsv() {
    $csvPath = Read-Host "Enter the CSV file path: "
    $users = Import-Csv -Path $csvPath

    foreach ($user in $users) {

        $userPrincipalName = $user.UserPrincipalName
        addOrRemoveFromCalendar $userPrincipalName
        }
    listUsers
}

function addOrRemoveASingleUser(){
    $userPrincipalName = Read-Host "Enter the username or email address: "
    
    addOrRemoveFromCalendar $userPrincipalName
    listUsers
}