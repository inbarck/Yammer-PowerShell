<#
    This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment.
    THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
    INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You
    a nonexclusive, royalty-free right to use and modify the Sample Code and to reproduce and distribute the object code form of
    the Sample Code, provided that You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which
    the Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded;
    and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’
    fees, that arise or result from the use or distribution of the Sample Code.
    Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within
    the Premier Customer Services Description.
#>

<####################################################################################################################################################

Yammer Module for Windows PowerShell v 1.0

 -------- USERS ---------------------------------------


#####################################################################################################################################################>



function Get-AllYammerUsersByPage([Int]$Page,[string]$BaseTarget) { 
        <#
        .DESCRIPTION
        This function gets page and base target and uses it to get Yammer users for that target.
        The function is used by Get-AllYammerUsers and Get-AllYammerUsersByGroupID 
        Internal function to be used only by Get-AllYammerUsers and Get-AllYammerUsersByGroupID 

        .EXAMPLE
        Get-AllYammerUsersByPage -Page 1 -BaseTarget 'https://www.yammer.com/api/v1/users.json?show_pending=true&page='

        .NOTES

        
        #>
    try {
            $target = $BaseTarget + $Page
            $response = Invoke-WebRequest -Uri $target -Headers $headers -Method GET 
            $users = ConvertFrom-Json $response.Content     
            Return $Users
       }    

    Catch [system.exception] {
      
        Return $_.Exception
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
    }

    
}


function Get-AllYammerUsers() { 
        <#
        .DESCRIPTION
        This function loop through users endpoint to get all the users in the network.
        It is using the function Get-AllYammerUsersByPage and sending the page number as long as it return 50 users. 
        Internal function to be used only by Get-YammerUser

        .EXAMPLE
        Get-AllYammerUsers

        .NOTES

        
        #>

    try {        
        
       $CurrentRequest = Get-AllYammerUsersByPage -Page 1 -BaseTarget 'https://www.yammer.com/api/v1/users.json?show_pending=true&page='
       $Page = 2
       $users= $CurrentRequest
       
       while ($CurrentRequest.Count -eq 50) {
            $CurrentRequest = Get-AllYammerUsersByPage -Page $Page -BaseTarget 'https://www.yammer.com/api/v1/users.json?show_pending=true&page='
            $users+= $CurrentRequest 
            $Page += 1 
       }
        Return $Users
   
    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $CurrentRequest
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
        Return $_.Exception
    }


}


function Get-AllYammerUsersByGroupID([Parameter(Mandatory=$true)][string]$groupID) { 
    <#
        .DESCRIPTION
        This function loop through /users/in_group endpoint for all users in a specific group.
        It is using the function Get-AllYammerUsersByPage and sending the page number as long as it return 50 users. 
        Internal function to be used only by Get-YammerUser

        .EXAMPLE
        Get-AllYammerUsers

        .NOTES

        
        #>
    
    try {        
            #getting the first page of users for the group
            $baseTarget = 'https://www.yammer.com/api/v1/users/in_group/' + $groupID + '.json?page='
            $CurrentRequest = Get-AllYammerUsersByPage -Page 1 -BaseTarget $baseTarget
            $Page = 2

            #checking if the request for users in a group was valid
            if ($CurrentRequest -like "*error*"){
                Write-Host $CurrentRequest
                Return $CurrentRequest
            }

            #Checking how many users were are in the group. If there are 0 users return an error message.

            elseif  ($CurrentRequest.users.Length -le 0){
                Write-Host 'There are no members in this group'
                Return $users
            }

            #if the request was valid we continue with the loop as long as the last request had 50 users
            else{
                
                $users= $CurrentRequest.users
                while ($CurrentRequest.users.Count -eq 50) {
                        $CurrentRequest = Get-AllYammerUsersByPage -Page $Page -BaseTarget $baseTarget
                        $users+= $CurrentRequest.users 
                        $Page += 1 
                        }
                 Return $Users

           }
        
    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $CurrentRequest
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
        Return $_.Exception
    }

    
}

function Get-YammerUserByEmail([Parameter(Mandatory=$true)][string]$email){
    <#
        .DESCRIPTION
        This function Return user according to their email address.
        The function is using the endpoint https://www.yammer.com/api/v1/users/by_email.json?email=
        Internal function to be used only by Get-YammerUser

        .EXAMPLE
        Get-YammerUserByEmail -email inbar@yeslogin.info

        .NOTES
   #>
    
    try{
            $target = 'https://www.yammer.com/api/v1/users/by_email.json?email=' + $email 
            $response = Invoke-WebRequest -Uri $target -Headers $headers -Method GET 
            $user =  ConvertFrom-Json $response.Content
            Return $user
            
    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
        Return $_.Exception
    }    
}

function Get-YammerUserByID([Parameter(Mandatory=$true)][string]$userID){
    <#
        .DESCRIPTION
        This function Return user according to their Yammer user ID.
        The function is using the endpoint https://www.yammer.com/api/v1/users/
        Internal function to be used only by Get-YammerUser

        .EXAMPLE
        Get-YammerUserByID -userID 1580914702

        .NOTES
   #>
    
    try{
        $target = 'https://www.yammer.com/api/v1/users/' +$userID + '.json'
        $response = Invoke-WebRequest -Uri $target -Headers $headers -Method GET 
        $user =  ConvertFrom-Json $response.Content
        Return $user
            
            #$user = Invoke-RestMethod $target -Headers $headers -Method GET
        
    }
    Catch [system.exception] {
        #Write-Host 'Error: Failed to execute request' $target
        #Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
        Return $_.Exception
    }
}


function Get-YammerCurrentUser(){
   <#
        .DESCRIPTION
        This function Return the current user - The user that we use their Token to retrieve.
        The function is using the endpoint https://www.yammer.com/api/v1/users/current.json
        Internal function to be used only by Get-YammerUser

        .EXAMPLE
        Get-YammerCurrentUser

        .NOTES
   #>
    
    try{
        $target = 'https://www.yammer.com/api/v1/users/current.json'
        $response = Invoke-WebRequest -Uri $target -Headers $headers -Method GET 
        $user =  ConvertFrom-Json $response.Content
        Return $user
    
       
        }
    
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
        Return $_.Exception
    }
    
}


function Get-YammerUser
(
        [string]$email,
        [string]$userID,
        [string]$groupID,
        [switch]$Currant
)
{
   <#
        .DESCRIPTION
        This is the main function to retrieve Yammer users.
        Users can be retrieved by Email, User ID, Group ID, Current 
        OR - if the function doesn't have any variable - it will retrieve all Yammer users (Only active and Pending).

        .EXAMPLE
        Get-YammerUser -email inbar@yeslogin.info
        Get-YammerUser -userID 1580914702
        Get-YammerUser -groupID 8740819
        Get-YammerUser -Currnt
        Get-YammerUser


        .NOTES
   #>



 try {
      #Get user by email address 
      If ($email) {$users = Get-YammerUserByEmail -email $email} 

      #Get user by User ID
      elseIf ($UserID) {$users =  Get-YammerUserByID -userID $userID}
               
      #Get Current user 
      elseif ($Currnt) {$users = Get-YammerCurrentUser}

      #Get all users that are member of Group by group ID
      elseIf ($groupID) {$users = Get-AllYammerUsersByGroupID -GroupID $groupID}

      #Get ALL users (when all variables are empty)
      elseIf ((!$email) -and (!$UserID) -and (!$GroupID) -and (!$Currnt)) {$users =  Get-AllYammerUsers}
    
    Return $Users            
        
    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
    }

    
}

<### C



###> 
function New-YammerUser
(
    [Parameter(Mandatory=$true)][string]$email,
    [string]$fullName,    
    [string]$jobTitle,
    [string]$departmentName,
    [string]$location,
    [string]$imProvider,
    [string]$imUsername,
    [string]$workTelephone,
    [string]$workExtension,
    [string]$mobileTelephone,
    [string]$externalProfiles,
    [string]$significantOther,
    [string]$kidsNames,
    [string]$interests,
    [string]$summary,
    [string]$expertise,
    [string[]]$education,
    [string[]]$previousCompanies

){ 
   <#
        .DESCRIPTION
            Create a new user in Yammer - Email is mandatory; Users are created as pending. 

            $education Format "school,degree,description,start_year,end_year"
            For example - "USC,MBA,Finance,2002,2004","UCLA,BS,Economics,1998,2002"

            $previousCompanies format "company,position,description,start_year,end_year"
            For example - "Geni.com,Engineer,my description,2005,2008"

        .EXAMPLE
            New-YammerUser -email 'User1@yeslogin.info' -fullName "User One" -jobTitle "Engineer" -departmentName "IT" -location "NYC" 
            -imProvider "12346578" -imUsername "User1" -workTelephone "987 654 1234" -workExtension "7984" -mobileTelephone "987 321 6544" 
            -externalProfiles "User1" -significantOther "User 2" -kidsNames "3 Users" -interests "Engineering" -summary "I'm an engineer" 
            -expertise "Computers" -education "USC,MBA,Finance,2002,2004","UCLA,BS,Economics,1998,2002" 
            -previousCompanies "Geni.com,Engineer,my description,2005,2008"


        .NOTES
   #>
    
    try {
        #$education=@("USC,MBA,Finance,2002,2004","UCLA,BS,Economics,1998,2002")
        $body = @{
            email = $email
            full_name = $fullName
            job_title = $jobTitle
            department_name = $departmentName
            location = $location
            im_provider = $imProvider
            im_username = $imUsername
            work_telephone = $workTelephone
            work_extension = $workExtension
            mobile_telephone = $mobileTelephone
            external_profiles = $externalProfiles
            significant_other = $significantOther
            kids_names = $kidsNames
            interests = $interests
            summary = $summary
            expertise = $expertise
            education = $education
            previous_companies = $previousCompanies
           
        } 


        $target = 'https://www.yammer.com/api/v1/users.json'
        $response = Invoke-WebRequest $target -Method POST -Body (ConvertTo-Json $body) -Headers $headers -ContentType 'application/json'
        $userURL= $response.Headers.Location
        $userID = $userURL -replace 'https://www.yammer.com/api/v1/users/'
        $user = Get-YammerUser -userID $userID 
        
        Return $user
        #>
    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
                          
   }
   if ($user) {Write-Host 'New User Created'} 
   
 
}


function Set-YammerUser
(
 [Parameter(Mandatory=$true)][string]$userID,
    [string]$fullName,    
    [string]$jobTitle,
    [string]$departmentName,
    [string]$location,
    [string]$imProvider,
    [string]$imUsername,
    [string]$workTelephone,
    [string]$workExtension,
    [string]$mobileTelephone,
    [string]$externalProfiles,
    [string]$significantOther,
    [string]$kidsNames,
    [string]$interests,
    [string]$summary,
    [string]$expertise,
    [Datetime]$birthday,
    [string[]]$education,
    [string[]]$previousCompanies
){ 

    try {
      <#
        .DESCRIPTION
            Update existing user in Yammer, according to the user ID. 

            $education Format "school,degree,description,start_year,end_year"
            For example - "USC,MBA,Finance,2002,2004","UCLA,BS,Economics,1998,2002"

            $previousCompanies format "company,position,description,start_year,end_year"
            For example - "Geni.com,Engineer,my description,2005,2008"

        .EXAMPLE
            Set-YammerUser -userID 1580914702 -fullName "User One" -jobTitle "Engineer" -departmentName "IT" -location "NYC" 
            -imProvider "12346578" -imUsername "User1" -workTelephone "987 654 1234" -workExtension "7984" -mobileTelephone "987 321 6544" 
            -externalProfiles "User1" -significantOther "User 2" -kidsNames "3 Users" -interests "Engineering" -summary "I'm an engineer" 
            -expertise "Computers" -education "USC,MBA,Finance,2002,2004","UCLA,BS,Economics,1998,2002" 
            -previousCompanies "Geni.com,Engineer,my description,2005,2008"


        .NOTES
   #>

        If ($birthday) {
          
          $day = (Get-Date $birthday).Day
          $month = (Get-Date $birthday).Month
          }
      

        $body = @{
            full_name = $fullName
            job_title = $jobTitle
            department_name = $departmentName
            location = $location
            im_provider = $imProvider
            im_username = $imUsername
            work_telephone = $workTelephone
            work_extension = $workExtension
            mobile_telephone = $mobileTelephone
            external_profiles = $externalProfiles
            significant_other = $significantOther
            kids_names = $kidsNames
            interests = $interests
            summary = $summary
            expertise = $expertise
            birth_date =  @{month = $day
                            day = $month}
            education = $education
            previous_companies = $previousCompanies
           
        } 

       
        $target = 'https://www.yammer.com/api/v1/users/' +$userID + '.json'
        $response = Invoke-WebRequest -Uri $target -Headers $headers -Method Put -Body (ConvertTo-Json $body) -ContentType 'application/json'
        If ($response.StatusCode -eq '200') {
            Write-Host 'User' $userID 'was updated'
        }
        
        $user =  ConvertFrom-Json $response.Content
        Return $user

       
              
    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
    }

}

<#Deleteing users - 

#>
function Remove-YammerUser([Parameter(Mandatory=$true)][string]$userID) { 
      <#
        .DESCRIPTION
            Delete Yammer user - behaviour similar to bulk update:
            Active users will be suspended. 
            Pending and Suspended users will be deleted; 

        .EXAMPLE
            


        .NOTES
   #>

    try {
        
        #checking that the user exist in Yammer
        $user = Get-YammerUser -userID $userID
        If ($user -like "*error*") {
            Return $user
        }
        else{
            $state = $user.state

            If ($state -eq 'active'){
                    $target = 'https://www.yammer.com/api/v1/users/' +$userID + '.json'
                    $response = Invoke-WebRequest $target -Headers $headers -Method DELETE   
                    Write-Host 'Active user' $userID 'is now'  ($user = Get-YammerUser -userID $userID).state            
            }
            elseif ($state -eq 'pending'){
                    $target = 'https://www.yammer.com/api/v1/users/' +$userID + '.json?delete=true'
                     $response = Invoke-WebRequest $target -Headers $headers -Method DELETE 
                     Write-Host 'Pending user' $userID  'is now' ($user = Get-YammerUser -userID $userID).state    

            }
            elseif ($state -eq 'suspended'){
                    $choice = ""
                        while ($choice -notmatch "[y|n]"){
                            $choice = read-host "The user" $userID "Is suspended. Deleted users cannot be restored. Do you want to continue? (Y/N)"
                        }
                        if ($choice -eq "y"){
                            $target = 'https://www.yammer.com/api/v1/users/' +$userID + '.json?delete=true'
                            $response = Invoke-WebRequest $target -Headers $headers -Method DELETE 
                            Write-Host "The user" $userID "was deleted"
                        }
                        else {write-host 'the operation was aborted'}      
            }
            elseif ($state -eq 'deleted'){
                write-host 'User' $userID 'is already deleted'
            }
        }
    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
    }
}

function Get-WebRequest(){
   
   try {

    Invoke-WebRequest -Uri 'https://www.yammer.com/api/v1/users/current.json' -Headers $headers -Method GET
    }
        Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'headers' $headers
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
}
}


Export-ModuleMember -Function Get-WebRequest
Export-ModuleMember -Function Get-YammerUser
Export-ModuleMember -Function New-YammerUser
Export-ModuleMember -Function Set-YammerUser
Export-ModuleMember -Function Remove-YammerUser
