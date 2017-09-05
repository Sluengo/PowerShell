#############      ACTIVE DIRECTORY PROVISIONING SCRIPT     #############
#############          Written By: Steven Luengo            #############

Function userOU {
   
       $ou = Read-Host "What is the user's OU? Enter the number next to the name:`n
            
        1) Accounting `n
        2) Administrator `n
        3) Breeder Advocates `n
        4) Breeder Compliance & Relations `n
        5) Customer Advocates `n
        6) Education & Development `n
        7) Executives `n
        8) HR `n
        9) Systems Engineering `n
        10) Legal `n
        11) Marketing `n
        12) Payment Processing `n
        13) Public Affairs `n
        14) Puppy Profile Approval`n
        15) Quality Assurance`n
        16) Reception `n
        17) Puppy Concierges `n
        18) Sales Operations `n
        19) Sales Management `n
        20) Travel `n
        21) Software Development`n
         ";
           

        switch($ou) {
           1 {$ou = "Accounting"}
           2 {$ou = "Administrator"}
           3 {$ou = "Breeder Advocates"}
           4 {$ou = "Breeder Compliance & Relations"}
           5 {$ou = "Customer Advocates"}
           6 {$ou = "Education & Development"}
           7 {$ou = "Executives"}
           8 {$ou = "HR"}
           9 {$ou = "Systems Engineering"};
           10 {$ou = "Legal"}
           11 {$ou = "Marketing"}
           12 {$ou = "Payment Processing"}
           13 {$ou = "Public Affairs"};
           14 {$ou = "Puppy Profile Approval"};
           15 {$ou = "Quality Assurance"};
           16 {$ou = "Reception"};
           17 {$ou = "Puppy Concierges"};
           18 {$ou = "Sales Operations"};
           19 {$ou = "Sales Management"};
           20 {$ou = "Travel"};
           21 {$ou = "Software Development"};
           default {"Test"};
         }
         
         return $ou;
}





Function userDepartment {
   
       $department = Read-Host "What is the user's department? Enter the number next to the name:`n
            
        1) Accounting `n
        2) Administrator `n
        3) Breeder Advocates `n
        4) Breeder Compliance & Relations `n
        5) Customer Advocates `n
        6) Education & Development `n
        7) Executives `n
        8) HR `n
        9) Systems Engineering `n
        10) Legal `n
        11) Marketing `n
        12) Payment Processing `n
        13) Public Affairs `n
        14) Profile Approval`n
        15) Quality Assurance`n
        16) Reception `n
        17) Puppy Concierges `n
        18) Sales Operation `n
        19) Sales Management `n
        20) Travel `n
        21) Software Development
         ";
           

        switch($department) {
           1 {$department = "Accounting"}
           2 {$department = "Administrator"}
           3 {$department = "Breeder Advocates"}
           4 {$department = "Breeder Compliance & Relations"}
           5 {$department = "Customer Advocates"}
           6 {$department = "Education & Development"}
           7 {$department = "Executives"}
           8 {$department = "HR"}
           9 {$department = "Systems Engineering"}
           10 {$department = "Legal"}
           11 {$department = "Marketing"}
           12 {$department = "Payment Processing"}
           13 {$department = "Public Affairs"};
           14 {$department = "Profile Approval"};
           15 {$department = "Quality Assurance"};
           16 {$department = "Reception"};
           17 {$department = "Puppy Concierges"};
           18 {$department = "Sales Operation"};
           19 {$department = "Sales Management"};
           20 {$department = "Travel"};
           21 {$department = "Software Development"};
           default {"Test"};
         }
         
         return $department;
}




Function userSecurityGroup {
        
    $securityGroup = Read-Host "What security group is the user in? Enter the number next to the name:

    1) Accounting Users `n
    2) Accounting Managers `n
    3) Breeder Placement Users `n
    4) Breeder Placement Managers `n
    5) Breeder Relations Users `n
    6) Breeder Relations Managers `n
    7) Customer Care Users `n
    8) Customer Care Managers `n
    9) E&D Users `n
    10) E&D Managers `n
    11) ExecutiveManagement `n
    12) HR Users `n
    13) HR Managers `n
    14) IT Users `n
    15) IT Managers `n
    16) Legal Users `n
    17) Legal Managers `n
    18) Marketing Users `n
    19) Marketing Managers `n
    20) Payment Processing Users`n
    21) Payment Processing Managers `n
    22) Public Affairs Managers `n
    23) Public Affairs Users `n
    24) Puppy Ads Users`n
    25) Puppy Ads Managers `n
    26) QA Users `n
    27) QA Managers `n
    28) Reception Users `n
    29) Reception Managers `n
    30) Sales Users `n
    31) Sales Managers `n
    32) Travel Users `n
    33) Travel Managers `n
    34) Web Users
    35) Web Managers

    ";

        
        switch($securityGroup) {
           1 {$securityGroup = "Accounting Users"}
           2 {$securityGroup = "Accounting Managers"}
           3 {$securityGroup = "Breeder Placement Users"}
           4 {$securityGroup = "Breeder Placement Managers"}
           5 {$securityGroup = "Breeder Relations Users"}
           6 {$securityGroup = "Breeder Relations Managers"}
           7 {$securityGroup = "Customer Care Users"}
           8 {$securityGroup = "Customer Care Managers"}
           9 {$securityGroup = "E&D Users"}
           10 {$securityGroup = "E&D Managers"}
           11 {$securityGroup = "ExecutiveManagement"}
           12 {$securityGroup = "HR Users"}
           13 {$securityGroup = "HR Managers"};
           14 {$securityGroup = "IT Users"};
           15 {$securityGroup = "IT Managers"};
           16 {$securityGroup = "Legal Users"};
           17 {$securityGroup = "Legal Managers"};
           18 {$securityGroup = "Marketing Users"};
           19 {$securityGroup = "Marketing Managers"};
           20 {$securityGroup = "Payment Processing Users"};
           21 {$securityGroup = "Payment Processing Managers"};
           22 {$securityGroup = "Public Affairs Managers"};
           23 {$securityGroup = "Public Affairs Users"};
           24 {$securityGroup = "Puppy Ads Users"};
           25 {$securityGroup = "Puppy Ads Managers"};
           26 {$securityGroup = "QA Users"};
           27 {$securityGroup = "QA Managers"};
           28 {$securityGroup = "Reception Users"};
           29 {$securityGroup = "Reception Managers"};
           30 {$securityGroup = "Sales Users"};
           31 {$securityGroup = "Sales Managers"};
           32 {$securityGroup = "Travel Users"};
           33 {$securityGroup = "Travel Managers"};
           34 {$securityGroup = "Web Users"};
           35 {$securityGroup = "Web Managers"};

           default {"Test"};
         }

    return $securityGroup;


}

Function firewallGroup {

    $firewallGroup = Read-Host "What firewall group should the user be in? Enter the number next to the name`n
 
    1) Basic `n
    2) Billing Users `n
    3) Partial Access `n
    4) Full Access `n
    5) Video Access `n
    
    ";

    switch($firewallGroup) {
        1 {$firewallGroup = "Basic"};
        2 {$firewallGroup = "Billing Users"};
        3 {$firewallGroup = "Partial Access"};
        4 {$firewallGroup = "Full Access"};
        5 {$firewallGroup = "Video Access"}

        default {"Test"};
    }

    return $firewallGroup;


}


Function userState {

    $state = Read-Host "What state is the user in? Enter the number next to the name`n
    
    1) FL`n
    2) CA`n
    3) UT`n
    ";

    switch($state) {
        1 {$state = "FL"};
        2 {$state = "CA"};
        3 {$state = "UT"};
        default {$state = "FL"};


    }

    return $state;


}


Function userOffice {

    $office = Read-Host "What office is the user in? Enter the number next to the name`n
    
    1) Florida`n
    2) California`n
    3) Utah`n
        ";

    switch($office) {
        1 {$office = "Florida"};
        2 {$office = "California"};
        3 {$office = "Utah"};
       
        default {$office = "Florida"};


    }

    return $office;


}






Function managerField($state,$securityGroup) {

$utManagers = @{"Breeder Placement Users" = "BPhillips"; "Breeder Placement Managers" = "GLiberman"; "Sales Users" = "BPhillips"; "Web Users" = "MMishkoff"; "Sales Managers" = "GLiberman" };

$flManagers = @{"Accounting Managers" = "JPowers"; "Accounting Users" = "LSeaman"; "Kmandell" = "Administrator"; "Breeder Placement Managers" = "GLiberman"; "Breeder Placement Users" = "BPhillips"; "Breeder Relations Managers" = "KRod"; "Breeder Relations Users" = "LHorner";
"Customer Care Managers" = "KRod"; "Customer Care Users" = "CPidcoe"; "E&D Managers" = "JKreinberg"; "E&D Users" = "SLedford"; "ExecutiveManagement" = "GLiberman";"HR Managers" = "JKreinberg"; "HR Users" = "HAbdulhafiz"; "IT Managers" = "SSmith"; "IT Users" = "KMandell";
"Marketing Users" = "JGreen"; "Marketing Managers" = "JGreen"; "Payment Processing Managers" = "JPowers"; "Payment Processing Users" = "KBaker"; "Public Affairs Managers" = "JKreinberg"; "Public Affairs Users" = "KRod"; "Puppy ads Managers" = "KRod"; "Puppy Ads Users" = "CRice";
"QA Managers" = "KRod"; "QA Users" = "DGandy"; "Reception Managers" = "KRod"; "Reception Users" = "CRice"; "Sales Users" = "RChristensen"; "Sales Managers" = "GLiberman"; "Travel Managers" = "KRod"; "Travel Users" = "JJellison"; "Web Managers" = "MMishkoff"; "Web Users" = "MMishkoff" };

$caManagers = @{"Marketing Users" = "JGreen"; "Marketing Managers" = "JGreen"; "Web Managers" = "MMishkoff"; "Web Users" = "MMishkoff"; "Accounting Users" = "LSeaman"; "Accounting Managers" = "JPowers";}

# Manager must be User Logon Name. 

Write-Host $state

Write-Host $securityGroup

    switch($state) {
    
        "UT" {$manager = $utManagers[$securityGroup]}
        "FL" {$manager = $flManagers[$securityGroup]}
        "CA" {$manager = $caManagers[$securityGroup]}
        default {$manager = "No Manager"}

    }

  
    Write-Host $manager
    return $manager;

    

}

################# PROGRAM BEGINS HERE ##################

# Imports the active directory cmdlets for use in script. It's like a function library
import-module activedirectory
New-PSDrive -Name FLDC01 -PSProvider ActiveDirectory -Server FLDC01.mypuppyspot.com -Root "//RootDSE/" -Scope Global
cd FLDC01:


# Get Credentials

Write-Output "Please enter in your AD credentials";
$creds = Get-Credential;

# Checks to see if user is running powershell as Administrator

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {Write-Warning " You do not have administrator rights
to run this script! `nPlease re-run this script as an Administrator!" 
Break}

# Write-Host outputs data to the screen, like cout
Write-Host "";
Write-Host "";
Write-Host "Welcome To the Active Directory Provisioning Script!";
Write-Host "";
Write-Host "";

# Do...While loops work the same way
do {
    # This section gets input from the user. 
    # Variables all start with $, like PHP
    # Read-Host collects data from user, like cin
    $first_name = Read-Host "What is the user's first name?`n";
    Write-Host "";
    $last_name = Read-Host "What is the user's last name?`n";
    Write-Host "";
    $first_initial = $first_name.substring(0,1); #substring parses out the first letter of the first_name
    Write-Host "";
   


   
   # Password Check Loop
    $loop_switch = $true;
        while($loop_switch -eq $true) {
              Write-Host "";
              $password = Read-Host "What password do you want?";
              Write-Host "";
              $password_confirm = Read-Host "Re-type password";
            if ($password -eq $password_confirm) {
                    # This code takes that password and converts it to a Secure string, hashes it, and forces it through. Need this in order to store passwords on AD, if not, it wont work
                    $password = $password | ConvertTo-SecureString -AsPlainText -Force;
                    $loop_switch = $false
                } else {
                    Write-Host "Incorrect password. Please try again";
                    $loop_switch = $true;
                    
                }
                
            }

    # This block ads users extension number, state,department, distroGroup, and security group

    $ext = Read-Host "What is the users extension number? Please enter only numbers.`n";

    Write-Host "";

    $state = userState;

    Write-Host "";

    $office = userOffice;

    Write-Host "";

    $ou = userOU;

    Write-Host "";

    $department = userDepartment;

   
    Write-Host "";

    $securityGroup = userSecurityGroup;

    Write-Host "";

    $firewallGroup = firewallGroup;

    Write-Host "";

    $title = Read-Host "What is this user's title?";

    Write-Host "";

    $manager = managerField $state $securityGroup; # passing variables to the function managerField. They must be in this format.


    
    #Test for existence of OU
    Try { 
            Get-ADOrganizationalUnit -Identity "OU=$ou,OU=$state,OU=PBBUsers,DC=PBB,DC=corp";
        } 
        
    Catch {
            
            Write-Output "OU does not exist. Creating new one..."

            New-ADOrganizationalUnit -Name $ou -path "OU=$state, OU=PBBUsers, DC=PBB, DC=corp";
          }


    



    # Concatenates first initial and last name for proper processing in New-ADUser, ex: takes "Steven" + "Luengo" and makes "Sluengo"
    $login_name = $first_initial + $last_name;
    $display_name = $first_name + " " + $last_name;
    $email = $first_initial + $last_name + "@puppyspot.com";
    $profile_path = "OU=$ou,OU=$state, OU=PBBUsers, DC=pbb, DC=corp"; #this sticks them in the right OU
    $server = "FLDC01.pbb.corp";

   



    # Logic block to determine which telephone number to put

    if ($department -eq "Puppy Concierges" -or $department -eq "Sales Management" -or $department -eq "Sales Operation") {
        
        $tele = "(866) 306-6064 ";
    
    } else {
    
        $tele = "(866) 706-7337 ";
    
    }


    # Logic block to determine address tab

     switch($state) {
        "FL" {$street = "7261 Sheridan Street, Suite 300A"; $city = "Hollywood"; $zipCode = "33024"};
        "CA" {$street = "9724 Washington Blvd."; $city = "Culver City"; $zipCode = "90232"};
        "UT" {$street = "2801 N. Thanksgiving Way, Suite 350"; $city = "Lehi"; $zipCode = "84043"};
        default {$street = "7261 Sheridan Street, Suite 300A"; $city = "Hollywood"; $zipCode = "33024"};


    }


    $company = "PuppySpot Group, LLC"

    # Normal syntax: New-ADUser -SamAccountName $login_name -Name $display_name -GivenName $first_name -Surname $last_name -DisplayName $display_name -UserPrincipalName $login_name -AccountPassword $password -Enabled $true -Path "OU=$department, OU=PBBUsers, DC=pbb, DC=corp";
    
    # This is a Hash Table, or Hash(name-value pair). The names on the left are the actual flags that the New-ADUser cmdlet uses. The values on right are variables which stored user input from before. Stored in the $user variable.
    # See comment above to see normal New-ADUser syntax 
    $user = @{
        SamAccountName = $login_name 
        Name = $display_name 
        GivenName = $first_name 
        Surname = $last_name 
        DisplayName = $display_name 
        EmailAddress = $email;
        UserPrincipalName = $email;
        AccountPassword = $password 
        Enabled = $true 
        Department = $department
        Office = $office;
        streetAddress = $street; 
        City = $city; 
        State = $state;
        postalCode = $zipCode;
        Manager = $manager;
        Title = $title;
        Company=$company;
        OfficePhone = $tele + " ext. " + $ext;
        Path = $profile_path;


    }
    
    # All that data gets passed in to New-ADUser using @user.  Must use the @ symbol not $ when passing a hash. 
    New-ADUser @user 

    # Adds SMTP proxy address to user.
    Set-ADUser $login_name -Add @{ProxyAddresses="SMTP:$email"};
    
    # Add-ADGroupMember is a cmdlet that assigns members in AD to AD groups. -Identity groupName -Members memberName
    # Add-ADGroupMember -Identity ReDir -Members $login_name;
    Add-ADGroupMember -Identity $securityGroup -Members $login_name;
    Add-ADGroupMember -Identity $firewallGroup -Members $login_name;
   
    # This section of code sets the PRIMARY Security group of the user. 
    $TokenPart1 = Get-ADGroup -Identity "Domain Users" -Properties primarygrouptoken | select primarygrouptoken;
    $Token = $TokenPart1 | Select -ExpandProperty primarygrouptoken;

    # DistinguishedName
    $DNName = "CN=" + $display_name + "," + $profile_path;

    #Replace variable

    $Replace = @{"PrimaryGroupID"="$Token"};

    # replace default PrimaryGroup
    Set-ADObject $DNName -Server $server -Replace $Replace;
    
    
    Write-Host "";
    Write-Host "";

    # Runs the SYNC cycle command on the primary DC
    Invoke-Command -ComputerName fldc01 -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta} -Credential $creds;
   
   # Logic block requesting whether you want to add another user or not, allows for multiple people to be added.
    $answer = Read-host "Do you want to add another user? Enter 1 for YES and 2 for NO";
        if ($answer -eq "1") {
            
            $answer = 1;
        } else {
            
            $answer = 2;
        }
    
    
    
} while($answer -eq 1);
