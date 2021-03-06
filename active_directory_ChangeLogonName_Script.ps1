#############      ACTIVE DIRECTORY RENAME USER LOGON NAME  #############
#############             Written By: Steven Luengo         #############

import-module activedirectory

$ou = Read-Host "What OU do you want to change?";

$usersToChange = Get-ADUser -Filter * -SearchBase "OU=$ou, OU=PBBUsers, DC=pbb, DC=corp";
# Grabs all the users in the Sales OU and puts them in an array

foreach ($user in $usersToChange) {

    # We can also substitute GivenName for sAmAccountName to get the user's first Name.
    $oldLogon = Get-AdUser $user -Properties * | Select-Object -Expand sAMAccountName;
    # This grabs Steven.L

    $lastName = Get-AdUser $user -Properties * | Select-Object -Expand surname;
    # This grabs Luengo

    $parse = $oldLogon.SubString(0,1);
    # This parses out "S"

    $newLogon = $parse + $lastName;
    # This concatenates "S" + "Luengo" into "SLuengo"
    
    $principalName = $newLogon + "@pbb.corp";
    
    
    
    Get-ADUser -Identity $user | Set-AdUser -Replace @{SamAccountName = $newLogon};
   
    Get-ADUser -Identity $user | Set-AdUser -Replace @{UserPrincipalName = $principalName};

    
    Set-AdUser -Identity $newLogon -HomeDrive "H:" -HomeDirectory "\\sv-fs-01\users$\$oldLogon"; 
    
   
    
    #Rename-Item -path "\\sv-fs-01\users$\$oldLogon" -newName "\\sv-fs-01\users$\$newLogon";
    # This renames the path of the home folder for the user to match the new logon name 
    
    
    #Set-AdUser -Identity $newLogon -HomeDrive "H:" -HomeDirectory "\\sv-fs-01\users$\$newLogon"; 
    # MUST APPLY THIS IN PROFILE IN AD
   
 
}


