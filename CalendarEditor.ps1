# To get the path where the ps1 files we want to get are stored
$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$scriptSubDirectoryFunctions = $scriptDirectory.ToString() + "\hooks"
$functionsFile = Join-Path -Path $scriptSubDirectoryFunctions -ChildPath "functions.ps1"
$bannerFile = Join-Path -Path $scriptSubDirectoryFunctions -ChildPath "banner.ps1"


# Dot source the file to import the functions
. $functionsFile
. $bannerFile

showBanner

##The user has to authenticate to the o365 admin account first
Connect-ExchangeOnline
## The user has to enter the calendar email address 
$calendarEmailAddress =  Read-Host "Enter the calendar email address: "
$calendarEmailAddress = $calendarEmailAddress + ":\Calendar"
#This input will allow the user to choose if he wants to Add, Remove or List the calendar members
$calendarAction = Read-Host "Enter if you want to Add members, Remove members or List the calendar (ADD\REMOVE\LIST): "



if($calendarAction -eq "list"){
    listUsers
}


else{

    $multipleUsersOrNot = Read-Host "Do you want to add users from a CSV file? (If not, then you will be able to add a single user)¨(YES\NO): "

    if($multipleUsersOrNot -eq "yes"){
       addOrRemoveUsersFromCsv
    }
    else{
       addOrRemoveASingleUser
    }

}


