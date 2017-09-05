# \\SV-FS-01\sip

# Checks to make sure the user has started Powershell as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {Write-Warning " You do not have administrator rights
to run this script! `nPlease re-run this script as an Administrator!" 
Break}


### FUNCTIONS ###

Function deleteMac($ext){

    Write-Output "Finding Files...`n"

    $files = Get-ChildItem -Path "\\RSAP01\sip\" | ? -FilterScript {$_.name -match "[A-Z0-9]{12}\.cfg"} | select-string -pattern "$ext"  | Select -ExpandProperty Filename


        if ($files -eq $null) {

            Write-Output "No Files were found.`n";
            return;

        } else {
    
            Write-Output "These are the files that we found:`n";
            Write-Output $files`n;
        }
    

    $decision = Read-Host "Would you like to delete them? Enter 1 for yes; Enter 2 for no.`n"

    if ($decision -eq "1") {
        Foreach ($file in $files) {
                
                Remove-Item \\RSAP01\sip\$file -Force;
                 
            }

        } Else {

            return; # Return back to the main Program

        }
    
    return; # Return back to the main Program

}




Function ModifyExt($name,$extension,$lineKeys) {
    # The backtick ` , is so you can write code without breaking the command.
    New-Item \\RSAP01\sip\x$extension.cfg -type file -force -Value " `
<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?> 
  <reginfo> 
	  <msg 
		  msg.bypassInstantMessage=""1""
		  msg.mwi.1.callBackMode=""contact""
		  msg.mwi.1.callBack=""899"" 
	   /> `
	   <reg `
		  reg.1.address=""$extension"" 
		  reg.1.auth.password=""Gandalf6219""
		  reg.1.auth.userId=""$extension""
		  reg.1.displayName=""$name""
		  reg.1.label=""$extension""
		  reg.1.lineKeys=""$lineKeys""
	   /> 
</reginfo> 


";
}


Function ModifyMac($macAddress,$extension) {
    New-Item \\RSAP01\sip\$macAddress.cfg -type file -force -Value "

<?xml version=""1.0"" standalone=""yes""?>
   <APPLICATION
	  APP_FILE_PATH=""sip.ld""
	  CONFIG_FILES=""x$extension.cfg, sip.cfg, features.cfg""
	  MISC_FILES=""""
	  LOG_FILE_DIRECTORY=""/log""
	  OVERRIDES_DIRECTORY=""/overrides""
	  CONTACTS_DIRECTORY=""/contacts""
/>

"

}




function FactoryResetPhone($oldExt) {


# Ask for extension to check against and also asks for Switchvox Admin Credentails. 
$extension = $oldExt;
$credential = Get-Credential

#Requests Switchvox site
$mainSite = Invoke-WebRequest "https://phone.mypuppyspot.com/admin" -SessionVariable session

# Grabs the forms on the page
$form = $mainSite.Forms[0];

$form.Fields["uid"] = $credentail.UserName;
$form.Fields["password"] = $credentail.Password;

# Log into Switchvox
$result = Invoke-WebRequest -uri ("phone.mypuppyspot.com/admin" + $form.Action) -WebSession $session -Method Post -Body $form.fields 

# Makes three xml requests to get all extensions. 
Invoke-RestMethod https://phone.mypuppyspot.com/xml  -WebSession $session  -Credential $credential -ContentType text/xml -Method Post -InFile C:\PScripts\xmlRequests\pageOne.xml -OutFile C:\PScripts\xmlResponses\phoneResponse.xml

$phoneFile = "C:\PScripts\xmlResponses\phoneResponse.xml";

  
# Retreives the xml content of each file.
[XML]$phoneConfig = Get-Content –Path $phoneFile;

# Selects the sip_phones child node
$phones = $phoneConfig | Select-Xml -XPath "//sip_phones" | Select -ExpandProperty "node"

# Variable used to hold phone object
$neededPhone = $phones | Select -ExpandProperty ChildNodes | Select extension, ip_address,state | Where {$_.extension -eq $extension}

$state = $neededPhone.state;
                 
    if ($state -eq "registered") {
                
        # Grab the IP address property from the phone object.
        $IPAddress = $neededPhone.ip_address;
                    
        Write-Output "This is the extension:" $neededPhone.extension
        Write-Output "This is the ip address:" $neededPhone.ip_address;
        Write-Output "This phone is currently:" $neededPhone.state;

        Write-Output "Factory Reseting Phone..."

        # Start a Internet Explorer and navigate to the IPAddress of the phone.
        $IE = new-object -ComObject internetexplorer.application
        $IE.navigate("https://$IPAddress/login.htm");
        $IE.visible = $true;

        # Allow time for IE to load.
        while ($IE.busy) {sleep 8}
        
        # Click Through the Warning Page due to certificate error.
        $link = $IE.Document.getElementById("overridelink");
        $link.click();

        while ($IE.busy) {sleep 8}


        # Enter password into password Field
        $passField = $IE.Document.getElementsByTagName('input') | Where {$_.getAttributeNode('class').Value -eq "loginField"}
        $passField.value = '456'

         while ($IE.busy) {sleep 8}
        # Click Submit to Login to Phone
        $submit = $IE.Document.getElementsByTagName('input') | Where {$_.getAttributeNode('class').Value -eq "button gray medium"}
        $submit.Click();

        while ($IE.busy) {sleep 5}



        if ($submit -ne $null) {
            # Enter password into password Field
            $passField = $IE.Document.getElementsByTagName('input') | Where {$_.getAttributeNode('class').Value -eq "loginField"}
            $passField.value = '159357'

            while ($IE.busy) {sleep 8}
             # Click Submit to Login to Phone
            $submit = $IE.Document.getElementsByTagName('input') | Where {$_.getAttributeNode('class').Value -eq "button gray medium"}
            $submit.Click();

        } 

        while ($IE.busy) {sleep 8}
        # Navigate to the Restore Section of the Site
        $IE.navigate("https://$IPAddress/phoneBackupRestore.htm");

        # Allow time for IE to load.
        while ($IE.busy) {sleep 8}

        # Click the Submit button to factory reset the phone
        $submit = $IE.Document.getElementsByTagName('input') | Where {$_.name -eq "RestoreToFacrotyBtn"}
        $submit.click();

        # Allow time for IE to load.
        while( $IE.busy) {sleep 8}

        # CLick Yes to confirm factory reset
        $confirmBtn = $IE.Document.getElementById("popupbtn0")
        $confirmBtn.click();

        while( $IE.busy) {sleep 8}

        $IE.Quit();
 }
   
   

}


Function ExtensionCheck {

        $oldExt = Read-Host "What is the extension you are trying to use? Just input numbers.`n";   
        
        deleteMac $oldExt;

        Write-Output "Checking to see if any phone is still connected..please input Switchvox Credentials.."

        FactoryResetPhone $oldExt;


}






Function SwitchvoxExtensionCreation($credential,$session) {


    do {

    
    
    $fName = Read-Host "What is the users first name?"
    $lName = Read-Host "What is the users last name?"
    $location = Read-Host "Where is the user located: FL, UT, or CA?"
    $extension = Read-Host "Please input the Extension? Only numbers"


    # FUNCTIONS

    function getID {
        Invoke-RestMethod https://phone.mypuppyspot.com/xml  -WebSession $session  -Credential $credential -ContentType text/xml -Method Post -InFile C:\PScripts\xmlRequests\getExt.xml -OutFile C:\PScripts\xmlResponses\getExtResponse.xml


        $getID = "C:\PScripts\xmlResponses\getExtResponse.xml";
 
        # Retreives the xml content of each file
        [XML]$getIDConfig = Get-Content –Path $getID;

        # Selects the sip_phones child node
        $id = $getIDConfig | Select-Xml -XPath "//extensions" | Select -ExpandProperty "node"

        $neededID = $id | Select -ExpandProperty ChildNodes | Select -ExpandProperty account_id

        return $neededID;

    }

    function getGroupID {

        $groupID = Read-Host "What Extension Group should this user be in? Enter the number next to the name:
    
        1) Accounting (FL)`n
        2) Breeder Advocates (FL) `n
        3) Breeder Advocates (UT) `n
        4) Customer Adv and Health (FL) `n
        5) IT `n
        6) Payment Processing (FL) `n
        7) Profile Approval (FL) `n
        8) Puppy Concierge (FL) `n
        9) Puppy Concierge (UT) `n
        10) Reception (FL) `n
        11) Sales Operations (FL) `n
        12) Sales Operations (UT) `n
        13) Travel (FL) `n
        ";

            switch($groupID) {
                1 {$groupID = "3751"}
                2 {$groupID = "1046"}
                3 {$groupID = "2695"}
                4 {$groupID = "1048"}
                5 {$groupID = "1049"}
                6 {$groupID = "1045"}
                7 {$groupID = "1050"}
                8 {$groupID = "2703"}
                9 {$groupID = "2704"}
                10 {$groupID = "1051"}
                11 {$groupID = "2706"}
                12 {$groupID = "2707"}
                13 {$groupID = "2705"};
      
                default {"Test"};
                }

        return $groupID;

    }



    # Creates file needed to retreive ID of extension
    New-Item "C:\PScripts\xmlRequests\getExt.xml" -type file -force -Value "<?xml version='1.0'?>
    <request method='switchvox.extensions.getInfo'>
	    <parameters>
		    <extensions>
			    <extension>$extension</extension>
		    </extensions> 
	    </parameters>
    </request>



    "


    #Requests Switchvox site
    $mainSite = Invoke-WebRequest "https://phone.mypuppyspot.com/admin" -SessionVariable session

    # Grabs the forms on the page
    $form = $mainSite.Forms[0];

    $form.Fields["uid"] = $credential.UserName;
    $form.Fields["password"] = $credential.Password;

    # Log into Switchvox
    $result = Invoke-WebRequest -uri ("phone.mypuppyspot.com/admin" + $form.Action) -WebSession $session -Method Post -Body $form.fields 


    # Calls the ID function to retrieve the ID needed to update the Extension
    $neededID = getID; 

    # Grab Group ID
    $neededGroupId = getGroupID;

    $sales = Read-Host "Is this a Puppy Conceirge Extension? Enter 1 for Yes. Enter 2 for No.";

    if ($sales -eq "1") {
        # Create POST XML file to be used to update the extension
        New-Item "C:\PScripts\xmlRequests\updateExt.xml" -type file -force -Value "<?xml version='1.0'?>
        <request method='switchvox.extensions.phones.sip.update'>
            <parameters>
                <account_id>$neededID</account_id>
                <first_name>$fName</first_name>
                <voicemail_password>$extension</voicemail_password>
                <last_name>$lName</last_name>
                <email_address>$email</email_address>
                <location>$location</location>
                <lang_locale>en_us</lang_locale>
                <extension_groups>
                    <extension_group>
                        <id>1</id>
                        <id>3599</id>
                        <id>$neededGroupID</id>
                    </extension_group>
                </extension_groups>
                </parameters>
        </request>

        "

    } else {


        # Create POST XML file to be used to update the extension
        New-Item "C:\PScripts\xmlRequests\updateExt.xml" -type file -force -Value "<?xml version='1.0'?>
        <request method='switchvox.extensions.phones.sip.update'>
            <parameters>
                <account_id>$neededID</account_id>
                <first_name>$fName</first_name>
                <voicemail_password>$extension</voicemail_password>
                <last_name>$lName</last_name>
                <email_address>$email</email_address>
                <location>$location</location>
                <lang_locale>en_us</lang_locale>
                <extension_groups>
                    <extension_group>
                        <id>1</id>
                        <id>1047</id>
                        <id>$neededGroupID</id>
                    </extension_group>
                </extension_groups>
                </parameters>
        </request>

        "

    }




    # Update Extension
    Invoke-RestMethod https://phone.mypuppyspot.com/xml  -WebSession $session  -Credential $credential -ContentType text/xml -Method Post -InFile C:\PScripts\xmlRequests\updateExt.xml -OutFile C:\PScripts\xmlResponses\newExt.xml

        
    
   



    # Logic block requesting whether you want to add another extension or not, allows for multiple people to be added.
    $answer = Read-host "Do you want to add another user to Switchvox? Enter 1 for YES and 2 for NO";
        if ($answer -eq "1") {
            
            $answer = 1;
        } else {
            
            $answer = 2;
        }

 } while ($answer -eq 1)



}

### END FUNCTIONS ###



### MAIN PROGRAM ####

do {

$switch = $true;

    while ($switch -eq $true) {

        $check = Read-Host "Have you run the extension Check? Enter 1 for NO. Enter 2 for YES.";

        if ($check -eq "1") { 
            
            ExtensionCheck;
        }
         
        $functionChoice = Read-Host "If you are modifying/creating an EXTENSION, please enter 1.`nIf you are modifying/creating a MAC ADDRESS, please enter 2.`n";

        if ($functionChoice -eq "1") {
            $name = Read-Host "Please input the display name for user`n";
            $extension = Read-host "Please input a 3 digit extension.`n";
            $lineKeys = Read-Host "Please input the number of lines(1 - 6 only).`n";
            # When passing parameters to functions, sepearte paramters by a space ONLY. NOTHING ELSE
            ModifyExt $name $extension $lineKeys;  
            $switch = $false;
        } 

        

        elseif($functionChoice -eq "2") {
            $extension = Read-host "Please input a 3 digit extension.`n";
            $macAddress = Read-host "Please input the mac address of the phone.`n";
            ModifyMac $macAddress $extension;  
            $switch = $false;
    
        }

       
        else {
    
            Write-Host "You have entered an invalid choice, please try again.`n" 


        }

    }

    # Logic block requesting whether you want to add another user or not, allows for multiple people to be added.
    $answer = Read-host "Do you want to add another file? Enter 1 for YES and 2 for NO";
        if ($answer -eq "1") {
            
            $answer = 1;
        } else {
            
            $answer = 2;
        }
    
} while ($answer -eq 1)



$switchvoxExt = Read-Host "Do you need to modify a user extension in Switchvox? Enter 1 for YES or 2 for NO."

if ($switchvoxExt -eq "1") {

    Write-Output "Please enter your switchvox credentials";

    $credential = Get-Credential;

    SwitchvoxExtensionCreation $credential $extension $session;

} else {

    exit;
}

### END MAIN PROGRAM ###