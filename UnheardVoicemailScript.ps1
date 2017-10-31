## Function List ##

# Creates a Voicemail XML Request File
Function CreateVMFile ($userID) {

      New-Item C:\PScripts\SalesRepVMs\xmlRequests\Voicemailrequest.xml -type file -force -Value "<?xml version=""1.0""?>
            <request method=""switchvox.users.voicemail.getList"">
                <parameters>
                    <account_id>$userID</account_id>
                    <folder>INBOX</folder>
                    <sort_field>date</sort_field>
                    <sort_order>ASC</sort_order>
                    <items_per_page>50</items_per_page>
                    <page_number>1</page_number>
                </parameters>
            </request>

    " | Out-Null;




}

# Creates an Extension XML Request File
Function CreateExtFile ($ext) {

      New-Item C:\PScripts\SalesRepVMs\xmlRequests\Extrequest.xml -type file -force -Value "<?xml version=""1.0""?>
 <request method=""switchvox.extensions.getInfo"">
	<parameters>
		<extensions>
			<extension>$ext</extension>
		</extensions>
	</parameters>
</request>
    " | Out-Null;




}

# Returns the number of unread Voicemails
Function GetVM($userID, $credential) {
    
   # Write-Host $userID;
    CreateVMFile $userID;
 

    
   
   # Makes three xml requests to get all extensions. 
   Invoke-RestMethod https://phone.mypuppyspot.com/xml  -Credential $credential -ContentType application/xml -Method Post -InFile C:\PScripts\SalesRepVMs\xmlRequests\Voicemailrequest.xml -OutFile C:\PScripts\SalesRepVMs\xmlResponse\Voicemailresponse.xml;


   $XMLResponseFilePath = "C:\PScripts\SalesRepVMs\xmlResponse\Voicemailresponse.xml";

    $xml = [ xml ](Get-Content $XMLResponseFilePath)
    $rows = $xml.SelectNodes("/response/result/messages/message") | Where-Object { $_.read -eq 'No' }
    $unread = @($rows).count
    $unread

    Return;


   
}

# Returns the sales Reps name
Function GetName ($ext, $credential) {

   
    CreateExtFile $ext;
 
    
   
    #$credential = Get-Credential;
     # Makes three xml requests to get all extensions. 
    Invoke-RestMethod https://phone.mypuppyspot.com/xml  -Credential $credential -ContentType application/xml -Method Post -InFile C:\PScripts\SalesRepVMs\xmlRequests\Extrequest.xml -OutFile C:\PScripts\SalesRepVMs\xmlResponse\Extresponse.xml;


    $extFile = "C:\PScripts\SalesRepVMs\xmlResponse\Extresponse.xml";

  
    # Retreives the xml content of each file.
    [XML]$extConfig = Get-Content –Path $extFile;



    # Selects the voicemail child node
    $parse = $extConfig | Select-Xml -XPath "//extensions" | Select -ExpandProperty "node";



    # Variable used to hold vm object
    $name = $parse | Select -ExpandProperty ChildNodes | select display;

    

    $newName = $name.display;
    
    $newName;

    Return; 



}

# Build Email
Function BuildHtml ($hash) {

## Building the HTML Page ##

$output = @"

<html>

<head>

<style>

BODY{background-color:peachpuff;width: 100%; font-size: 1em}
H2{font-size: 120%}
H4{}
TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;margin-left:auto;margin-right:auto;}
TH{font-size: 120%; border-width: 1px;padding: 8px;border-style: solid;border-color: black; background-color:#63d14f}
TD{font-size: 105%; text-align: center;border-width: 1px;padding: 4px;border-style: solid;border-color: black;background-color:white}

</style>

</head>

<body>

<h2> Hello,<br><br> Please see the list of reps who have un-heard voicemails below.</h2>
<h3> Thank you, <br> Systems Engineering</h3>

<h3><u> Number of Unread Voicemails</u></h3>

<table>`r`n<tr><th>Name</th><th>Number of Voicemails</th></tr>`r`n
"@

$hash.GetEnumerator() | % {
    $output += "<tr><td>$($psitem.key)</td><td>$($psitem.value)</td></tr>`r`n"
}
$output += "</table></body></html>" |Out-File C:\PScripts\SalesRepVMs\html\htmlTest.html 

 

Return $output

}


## Sends Email
Function SendEmail ($output, $department) {

    
    Write-Host "In Sendmail";
    
    Write-Host $department;
 
    # $department contains the $increment value. Returns the appropriate subject and To field depending on the number.
    # Order is important here, since Accounting is the first department in $repArrayOfHashes, it must be number 1   
    switch($department) {

        1 {$subject = "Accounting Unread Voicemail Report"; $To="sluengo@puppyspot.com"};
        2 {$subject = "Breeder Advocates and Research Unread Voicemail Report"; $To="sluengo@puppyspot.com"};
        3 {$subject = "Breeder Compliance and Relations Unread Voicemail Report"; $To="sluengo@puppyspot.com"};
        4 {$subject = "Customer Advocates Unread Voicemail Report"; $To="sluengo@puppyspot.com"};
        5 {$subject = "Education & Development Unread Voicemail Report"; $To="sluengo@puppyspot.com"};
        6 {$subject = "Payment Processing Unread Voicemail Report"; $To="sluengo@puppyspot.com"};
        7 {$subject = "Public Affairs Unread Voicemail Report"; $To="sluengo@puppyspot.com"};
        8 {$subject = "Profile Approval Unread Voicemail Report"; $To="sluengo@puppyspot.com"};
        9 {$subject = "Sales Unread Voicemail Report"; $To="sluengo@puppyspot.com"};
       10 {$subject = "Travel Unread Voicemail Report"; $To="sluengo@puppyspot.com"};
         
        default {$subject = "Unread Voicemail Report"; $To="sluengo@puppyspot.com"};


    }
 




 # Builds email message
$messageParameters = @{                        
                Subject = $subject
                                       
                Body = $output;                  
                                          
                From = "UnreadVoicemails@puppyspot.com"                        
                To = $To                        
                SmtpServer = "puppyspot-com.mail.protection.outlook.com"                        
            }



# Sends Email
Send-MailMessage @messageParameters -BodyAsHtml -UseSsl -port 25 



}

## List of of Department Specific Functions that retrieve un-heard voicemails

# Builds Hashes by pulling data from the phone system
Function BuildHash ($group_id, $credential) {

 Write-Host "In BuildHash Function";

 New-Item C:\PScripts\SalesRepVMs\xmlRequests\testThing.xml -type file -force -Value "<?xml version=""1.0""?>
<request method=""switchvox.extensionGroups.getInfo"">
	<parameters>
		<extension_group_id>$group_id</extension_group_id>    
	</parameters>
</request>

    " | Out-Null;



    Invoke-RestMethod https://phone.mypuppyspot.com/xml  -Credential $credential -ContentType application/xml -Method Post -InFile C:\PScripts\SalesRepVMs\xmlRequests\testThing.xml -OutFile C:\PScripts\SalesRepVMs\xmlResponse\testThingResponse.xml;


    $extFile = "C:\PScripts\SalesRepVMs\xmlResponse\testThingResponse.xml";

  
   $xml = [ xml ](Get-Content $extFile)
   $rows = $xml.SelectNodes("/response/result/extension_group/members")


   $rows = $rows | Select -ExpandProperty ChildNodes | Select extension, account_id 

   $rows | Export-Csv C:\PScripts\SalesRepVMs\csv\test.csv
  
   $table = Import-Csv C:\PScripts\SalesRepVMs\csv\test.csv

   $hashtable = @{};

   foreach($set in $table) {
        $hashtable[$set.extension]=$set.account_id
  }
    


    return $hashtable;





}



Function GetVoicemail () {

# Initialize all departmental hashes, calls BuildHash "groupID" phoneSystemCredential

$accountingReps = BuildHash "3751" $credential;

$FloridaBreederAdvocates = BuildHash "1046" $credential
$UtahBreederAdvocates = BuildHash "2695" $credential
$BreederResearch = @{"522"="1126";"453"="1182";"509"="1137";}

# Add the hashes together
$BreederAdvocates = $FloridaBreederAdvocates + $UtahBreederAdvocates + $BreederResearch;


$breederRelations = BuildHash "6267" $credential;

$FloridaCustomerAdvocates = BuildHash "1048" $credential
$UtahCustomerAdvocate = BuildHash "6108" $credential

$customerAdvocates = $FloridaCustomerAdvocates + $UtahBreederAdvocates;

$educationAndDevelopment = BuildHash "6269" $credential;

$hr = BuildHash "6268" $credential;

$paymentProcessing = BuildHash "1045" $credential;

$publicAffairs = BuildHash "6270" $credential;

$profileApproval = BuildHash "1050" $credential;

$reception = BuildHash "1051" $credential;

$FloridaSalesReps = BuildHash "2703" $credential;
$FloridaSalesOps = BuildHash "2706" $credential;
$UtahSalesReps = BuildHash "2704" $credential;
$UtahSalesOps = BuildHash "2707" $credential;

$salesReps = $FloridaSalesReps + $FloridaSalesOps + $UtahSalesReps + $UtahSalesOps;

$travel = BuildHash "2705" $credential;

# Initialize the peopleHash
$peopleHash = @{};


# Initialize the array that contains all the hashes
$repArrayOfHashes =@($accountingReps,$BreederAdvocates,$breederRelations,$customerAdvocates,$educationAndDevelopment,
$paymentProcessing,$publicAffairs,$profileApproval, $salesReps, $travel);



Write-Host "In GetVoiceMail Function";

# Start the increment variable to determine which department we are working with. Order is important.
$increment = 1;

# loops over the array of hashes 
Foreach ($hash in $repArrayOfHashes) {

  Write-Host "In GetVoiceMail Function";
   
    # Loops over the contents of each hash
    Foreach ($person in $hash.GetEnumerator()) {

        # Returns the number of unheard voicemails for the user
        $numOfVoicemails = GetVM $($person.Value) $credential;
        
            
        # Returns the name of the person associated with those voicemails
        $nameOfPerson = GetName $($person.Name) $credential;
        
        
        # Adds both values to the hash. Must convert string num of voicemails to integers
        $peopleHash.Add("$nameOfPerson",[int]$numOfVoicemails);

   
    }

    

    # Sorts hash by the number of unheard voicemails, most voicemails on top
    $peopleHash = $peopleHash.GetEnumerator() | sort -Descending -Property Value;

    # Calls the BuildHtml function that..builds the HTML file using the hash and returns the html code
    $output = BuildHtml $peopleHash
 
    # Calls SendEmail Function, sending the html code and the current increment( number corresponds with department).
    SendEmail $output $increment

    # Re-initializes the hash
    $peopleHash = @{};

    # Increments the value by 1 for the next department.
    $increment  = $increment +1;
}






}


# Main Program

Function Main () {
### BEGIN MAIN ###

$uname = "admin";
$pw    = get-content  C:\PScripts\strings\salesvmstring.txt | convertto-securestring
$cred = New-Object -typename System.Management.Automation.PSCredential -ArgumentList $uname,$pw;

$credential = $cred

GetVoicemail($credential);


}

## Start's the program

Main;