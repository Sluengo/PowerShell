################################### Offboarding Script, Written By: Steven Luengo -- Date: 8/15/2016 ######################################

# Imports AD module, connects to FLDC01, cds to that new PSDrive
Import-Module ActiveDirectory
New-PSDrive -Name FLDC01 -PSProvider ActiveDirectory -Server FLDC01.mypuppyspot.com -Root "//RootDSE/" -Scope Global
cd FLDC01:

Write-Output "Please input your AD admin credentials";

$creds = Get-Credential;

# Checks to see if user is running powershell as Administrator

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {Write-Warning " You do not have administrator rights
to run this script! `nPlease re-run this script as an Administrator!" 
Break}

# Checks to see if user is connected to an exchange session

If (!(Get-PSSession | Where {$_.ConfigurationName -eq "Microsoft.Exchange"})) {

    Write-Warning "You are not connected to Exchange. Please connect to an exchange session and re-run the script";
    Break;
}

###################################################



Function Remove-PBBUser {

$logfile = 'C:\PScripts\disableUserErrorLog';

$emails = Read-Host "What is the email address of the user(s) you are removing? Ex: Steven.L@puppyspot.com,JMandell@puppyspot.com.`n Please separate by a comma with no spaces`n" ;

$emailArray = $emails.split(',');

    foreach ($email in $emailArray) {

        $mobile = Get-MobileDevice -Mailbox $email | Select -ExpandProperty Identity;

        if($mobile -ne $null) {

            Try {

                $everthing_good = $true;
                # Removes OWA
                Set-CASMailbox -Identity $email -OWAforDevicesEnabled $False;
     

                } Catch {
            
                    $everything_good = $False;
                    Write-Warning "OWA disabling failed for $email"
                    $email | Out-File $logfile -Append
                    Write-Warning "Logged to $logfile";


                }
            



              if ($everthing_good) {
                 
                 Write-Output "OWA has been disabled for $email";
            
                # Removes ActiveSync
                 Set-CASMailbox -Identity $email -ActiveSyncEnabled $False;
                 Write-Output "ActiveSync has been disabled for $email";

                 # Removes Mobile Devices
                 # Get-MobileDevice -Mailbox $email | Remove-MobileDevice -Confirm $False;
                 # Write-Output "All mobile devices have been removed for $email";

               
                # Gets user's ad account using email address.  
                $username = Get-ADUser -filter{emailaddress -like $email} | Select -ExpandProperty SamAccountName;

                # Sets user's primary group to Domain Users

                $primaryGroup = Get-ADGroup "Domain Users" -Properties @("primaryGroupToken");
                Get-ADUser $username | Set-ADUser -Replace @{primaryGroupID=$primaryGroup.primaryGroupToken};

               
                # Output group membership

                $groups = Get-ADUser $username -Properties memberOf | Select -ExpandProperty memberOf;


                Write-Output "$username is a member of the following groups:"
                Write-Output "";
                Get-ADUser $username -Properties memberOf | Select -ExpandProperty MemberOf | Get-ADGroup -Properties name | Select Name | Sort Name;

                Write-Output "";

                # Removes User from groups

                Try {
                    $group_removal = $true;
                    $ADGroups = Get-ADPrincipalGroupMembership -Identity $username | where {$_.Name -ne "Domain Users"}
                    Remove-ADPrincipalGroupMembership -Identity $username -MemberOf $ADGroups -Confirm:$false;
                
                
                } Catch {
                    $group_removal = $false;
                    Write-Warning "$username was not removed from any groups";
                
                }
                    
                    if($group_removal -eq $true) {
                    
                        Write-Output "$username has been removed from all groups except Domain Users.";
                    
                    }

                    

                 # Changes Users Password
                Try {
                    $password_changed = $true;
                    Set-ADAccountPassword -Identity $username -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "P@ssw0rd5722!" -Force);
                   
                } Catch {
                    $password_changed = $false;
                    Write-Warning "$username password failed."
                }

                    if($password_changed -eq $true) {
                
                         Write-Output "$username password has been changed";
                    }

                # Disables AD Account. 
                Try {
                    $account_disabled = $true;
                    Disable-ADAccount -Identity $username;

                } Catch {
                    $account_disabled = $false;
                    Write-Warning "$username's account failed disabling.";
                }

                    if($account_disabled -eq $true) {
                
                         Write-Output "$username account has been disabled.";
                    }
                

                # Timestamps when the account was disabled to the description field

                Try {
                    $change_date = $true;
                    $date = Get-Date -format u;
                    Set-ADUser $username -Description $date;
                } Catch {
                    $change_date = $false
                    Write-Warning "$username's date field failed to change";
                }

                    if($change_date -eq $true) {
                        Write-Output "$username has been disabled on $date";
                    }

                # Moves AD account to Ex-Employee Distro

                Try {
                    $move_account = $true;
                    Search-ADAccount -AccountDisabled | Move-ADObject -TargetPath 'OU=Disabled Users,DC=pbb,DC=corp' 
                } Catch {
                    $move_account = $false;
                    Write-Warning "$username failed to move to Disabled Users OU";
                }
                    
                    if($move_account -eq $true) {
                        Write-Output "$username has been moved to the Disabled Users OU";
                    }
                
            }

        
        } else { # Defaults to this if ONLY if there are no mobile devices.
        
            
            $username = Get-ADUser -filter{emailaddress -like $email} | Select -ExpandProperty SamAccountName;


            # Sets user's primary group to Domain Users

            $primaryGroup = Get-ADGroup "Domain Users" -Properties @("primaryGroupToken");
            Get-ADUser $username | Set-ADUser -Replace @{primaryGroupID=$primaryGroup.primaryGroupToken};

               
            # Output group membership

            $groups = Get-ADUser $username -Properties memberOf | Select -ExpandProperty memberOf;


          
            Write-Output "$username is a member of the following groups:";

            Get-ADUser $username -Properties memberOf | Select -ExpandProperty MemberOf | Get-ADGroup -Properties name | Select Name | Sort Name;

            Write-Output "";

            # Removes User from groups

            Try {
                  $group_removal = $true;
                  $ADGroups = Get-ADPrincipalGroupMembership -Identity $username | where {$_.Name -ne "Domain Users"}
                  Remove-ADPrincipalGroupMembership -Identity $username -MemberOf $ADGroups -Confirm:$false;
                
                
                } Catch {
                    $group_removal = $false;
                    Write-Warning "$username was not removed from any groups";
                
                }
                    
                    if($group_removal -eq $true) {
                    
                        Write-Output "$username has been removed from all groups except Domain Users.";
                    
                    }



              # Removes OWA 

             Try {
                $owaRemove = $true;
                Set-CASMailbox -Identity $email -OWAforDevicesEnabled $False;
             

                } Catch {
                    
                    $owaRemove = $false;
                    Write-Warning "OWA disable for $email FAILED.";

                }

                    if ($owaRemove -eq $true) {
                    
                       Write-Output "OWA has been disabled for $email";
                    
                    }
    
              #####################################


              # Removes Active Sync

             Try {
                $activeSyncStatus = $true;
                Set-CASMailbox -Identity $email -ActiveSyncEnabled $False;
             

                } Catch {
                    
                    $activeSyncStatus = $false;
                    Write-Warning "ActiveSync disable for $email FAILED.";

                }

                    if ($activeSyncStatus -eq $true) {
                    
                       Write-Output "ActiveSync has been disabled for $email";
                    
                    }


               #####################################


             Try {
                    $password_changed = $true;
                    Set-ADAccountPassword -Identity $username -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "P@ssw0rd5722!" -Force);
                   
                } Catch {
                    $password_changed = $false;
                    Write-Warning "$username password failed."
                }

                    if($password_changed -eq $true) {
                
                         Write-Output "$username password has been changed";
                    }

            #####################################

            Try {
                    $account_disabled = $true;
                    Disable-ADAccount -Identity $username;

                } Catch {
                    $account_disabled = $false;
                    Write-Warning "$username's account failed disabling.";
                }

                    if($account_disabled -eq $true) {
                
                         Write-Output "$username account has been disabled.";
                    }
                

            #####################################

             Try {
                    $change_date = $true;
                    $date = Get-Date -format u;
                    Set-ADUser $username -Description $date;
                } Catch {
                    $change_date = $false
                    Write-Warning "$username's date field failed to change";
                }

                    if($change_date -eq $true) {
                        Write-Output "$username has been disabled on $date";
                    }

            #####################################

              Try {
                    $move_account = $true;
                    Search-ADAccount -AccountDisabled | Move-ADObject -TargetPath 'OU=Disabled Users,DC=pbb,DC=corp' 
                } Catch {
                    $move_account = $false;
                    Write-Warning "$username failed to move to Disabled Users OU";
                }
                    
                    if($move_account -eq $true) {
                        Write-Output "$username has been moved to the Disabled Users OU";
                    }
        
          }
         

    } # End of Foreach


  Invoke-Command -ComputerName fldc01 -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta} -Credential $creds;


} # End of Function





do {



Remove-PBBUser;




$answer = Read-host "Do you want to continue using the Program? Enter 1 for YES and 2 for NO";
          if ($answer -eq "1") {

              $answer = 1;

            } else {
            
              $answer = 2;
            }




} while($answer -eq 1);