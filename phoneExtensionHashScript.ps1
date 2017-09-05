################## Phone Extension AD Script ##################
################## Description: This script creates extensions for users in AD ##################
################## Written By: Steven Luengo ##################

#Importing csv file that contains names and extensions
$myTable = Import-Csv -Path \\sv-fs-01\it\userSuites.csv;

#creating an empty Hash
$hashTable = @{};

#Loop that populates hash with key,value pairs
foreach($set in $myTable) {
# ID in brakets takes cell inside ID column and sets it as key = Ext which is second column.
    $hashTable[$set.Username]=$set.Extension;

}

# This array mimics retrieving an array of names from AD
$nameArray = Get-ADUser -Filter * | Select -ExpandProperty samAccountName;


# loops through the name array looking for a name that matches a key in the hash
foreach($name in $nameArray) { 
    $name = $name.trim();
# If the hash contains a key with the name in the array
    if($hashTable.ContainsKey($name)) {
       
       # It will Grab the value associated with that key(in this case an extension) 
        $ext = $hashTable.Item($name)


       # This is where you can write code to Set-ADUser -Identity $name -OfficePhone $ext
       Set-ADuser -Identity $name -OfficePhone $ext;

        
    }

}


#$hashTable.GetEnumerator() | % {$_.Value};