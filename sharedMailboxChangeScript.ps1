Get-Mailbox -RecipientTypeDetails SharedMailbox |Select DisplayName, Emailaddresses | Export-Csv sharedSMTPaddressesBefore.csv

$array = Get-Mailbox -RecipientTypeDetails SharedMailbox  | Select -ExpandProperty DisplayName;

Write-Output "This is the array we are looping through"

Write-Output "";

Write-Output $array;

Write-Output "";

Write-Output "------Starting First Loop------"

foreach($displayName in $array) { # Loops through each name in the array 
    

    Write-Output "------You are in the first loop------"
    Write-Output "DisplayName:$displayName";

    Write-Output "";

    # Grabs the primary SMTP address for the current email and stores it.
    $primarySMTP = Get-Mailbox -RecipientTypeDetails SharedMailbox $displayName | select -ExpandProperty PrimarySmtpAddress;
    Write-Output "Primary SMTP:$primarySMTP";

    Write-Output "";
   
    # Grabs the actual email object
    $emailObject = Get-Mailbox -RecipientTypeDetails SharedMailbox $displayName
    Write-Output "Email Object:$emailObject";

    Write-Output "";

    # Uses the email object above to grab all the current email aliases
    $distroArray = $emailObject.EmailAddresses;

    Write-Output "These are the Aliases";
    Write-Output $distroArray;
    
    
    Write-Output "";


    foreach ($email in $distroArray) {

        Write-Output "------ You are in the second Loop------"
        Write-Output $email;
        
        # Grabs the position that the special character is in --> this case it is the @ symbol, stores its number in $index
        $index = $email.IndexOf("@");

        # Grabs string from the beginning all the way to special character(not including special character).
        $name = $email.Substring(0,$index);

    
         # Grabs string from the special character(not including speciail character) all the way to the end of the string. 
        $domain = $email.Substring($index+1);

        # Logic block that checks the domain and appends the new address to those accounts. 
    

        if ($domain -ne "puppyspot.com") { # Condition that adds @puppyspot alias to each email alias name
      
          $emailObject.EmailAddresses += "$name@puppyspot.com"; 
          $emailObject.EmailAddresses += "SMTP:$primarySMTP";

         
          }

          $noRepeats = $emailObject.EmailAddresses | sort -Unique;

          Write-output "This will be appended to $displayName"
          Write-Output ""
          Write-Output $noRepeats;

          Set-Mailbox -Identity $displayName -EmailAddresses $noRepeats;

          

    }





    Write-Output " ------Starting main loop over------ ";
    Write-Output "";







}

Get-Mailbox -RecipientTypeDetails SharedMailbox | Select DisplayName, EmailAddresses| Export-Csv sharedSMTPaddressesAfter.csv