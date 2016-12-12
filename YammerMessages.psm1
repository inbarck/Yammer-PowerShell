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

<#####################################################################################################################################################

Yammer Module for Windows PowerShell v 1.0

 -------- MESSAGES ---------------------------------------


######################################################################################################################################################>



function Get-YammerMessageByID([string]$messageID){
      <#
        .DESCRIPTION
         This function return message by message ID
         It is using the https://www.yammer.com/api/v1/messages.json endpoint
  
         Used by Get-YammerMessage.

        .EXAMPLE
        Get-YammerMessageByID -messageID 743074149 

        .NOTES

        
        #>
     try {
            $target = 'https://www.yammer.com/api/v1/messages/' + $messageID  +'.json'            
            $response = Invoke-RestMethod $target -Headers $headers -Method GET
           Return $response 
     
        }
        Catch [system.exception] {
            Write-Host 'Error: Failed to execute request' $target
            Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
          
        }
    
    
}

function Get-YammerMessages([string]$olderThan,[string]$newerThan,[boolean]$threaded,[string]$limit){
     <#
        .DESCRIPTION
        The following function is using messages endpoint to retrieve multiple messages, based on the provided parameters.
        https://www.yammer.com/api/v1/messages.json 

        Allowed parameters:
        * older_than - Returns messages older than the message ID specified as a numeric string
        * newerThan - Returns messages newer than the message ID specified as a numeric string
        * threaded - threaded=true will only return the first message in each thread. 
        threaded=false will return the thread starter messages in order of most recently active as well as the two most recent messages, as they are viewed in the default view on the Yammer web interface.
        * limit - Return only the specified number of messages.

        Used by Get-YammerMessage.

        .EXAMPLE
        Get-YammerMessageByID -messageID 743074149 

        .NOTES

        
        #>
try {
        $baseTarget = 'https://www.yammer.com/api/v1/messages.json'
        $parms = '?'
        if ($olderThan) {$parms += 'older_than=' + $olderThan + '&'}
        if ($newerThan) {$parms += 'newer_than=' + $newerThan + '&'}
        if ($threaded) {$parms += 'threaded=' + $threaded + '&'}
        if ($limit) {$parms += 'limit=' + $limit}
        
        ## Checking if the parms section ends with unwanted char like & or ? and removes it.
        if (($parms.Substring($parms.Length-1) -eq '?') -or (($parms.Substring($parms.Length-1) -eq '&'))){
            $parms = $parms.Substring(0,$parms.Length-1)
        } 

        #Base URL + Parms create a new target 
        $target = $baseTarget + $parms
        $target

        $response = Invoke-RestMethod $target -Headers $headers -Method GET
     
    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
    }
    
    Return $response.messages

}

function Get-YammerMessagesByThreadID([string]$threadID){
  <#
        .DESCRIPTION
        The following function return all messages according to thread ID.
        Since every call return only 20 messages, the function loop through thread messages until older_available is false.
        
        Used by Get-YammerMessage.

        .EXAMPLE
        Get-YammerMessagesByThreadID -threadID 744466867 

        .NOTES
        
        #>

    try {
            #base target
            $target =  'https://www.yammer.com/api/v1/messages/in_thread/' + $threadID  +'.json'            
            $response = Invoke-RestMethod $target -Headers $headers -Method GET
            $messages = $response.messages 
            
            #checking if additional messages in thread are available   
            $olderAvailable = $response.meta.older_available        

            while ($olderAvailable){
                    #Checking the oldest message ID
                    $ArrLength = ($messages.Length) -1
                    $oldestMessage = $messages[$ArrLength].id

                    #Creating new target base on the oldest message
                    $newTarget = $target + '?older_than=' + $oldestMessage
                    $response = Invoke-RestMethod $newTarget -Headers $headers -Method GET

                    #adding current response to the messages array 
                    $messages += $response.messages
                
                    #updating Oldest available 
                    $olderAvailable = $response.meta.older_available
                
            }
            
        }
        Catch [system.exception] {
            Write-Host 'Error: Failed to execute request' $target
            Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
        }
        Return $messages 
}


function Get-YammerMessagesByGroupID([string]$groupID){
  <#
        .DESCRIPTION
        The following function return all messages according to Group ID.
        Since every call return only 20 messages, the function loop through thread messages until older_available is false.
        
        Used by Get-YammerMessage.
        
        .EXAMPLE
         Get-YammerMessagesByGroupID -groupID 8740819 

        .NOTES

        
        #>

 try {
            #base target
            $target =  'https://www.yammer.com/api/v1/messages/in_group/' + $groupID  +'.json'            
            $response = Invoke-RestMethod $target -Headers $headers -Method GET
            $messages = $response.messages 
             
            #checking if additional messages in thread are available   
            $olderAvailable = $response.meta.older_available        

            while ($olderAvailable){
                    #Checking the oldest message ID
                    $ArrLength = ($messages.Length) -1
                    $oldestMessage = $messages[$ArrLength].id

                    #Creating new target base on the oldest message
                    $newTarget = $target + '?older_than=' + $oldestMessage
                    $response = Invoke-RestMethod $newTarget -Headers $headers -Method GET
                                     
                    #adding current response to the messages array 
                    $messages += $response.messages
                
                    #updating Oldest available 
                    $olderAvailable = $response.meta.older_available
            }
            
        }
        Catch [system.exception] {
            Write-Host 'Error: Failed to execute request' $target
            Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
        }
        Return $messages 
  

}

function Get-YammerMessagesByPage([Int]$page){
  <#
        .DESCRIPTION
        The following function return all messages according to page.
                
        Since every call return only 20 messages, the function loop through messages until until we get to the requested page. 
        The function also uses older_available.

        Used by Get-YammerMessage.
       
        .EXAMPLE
         Get-YammerMessagesByPage -page 12

        .NOTES

        #>

  try {
        #First request
        $response = Invoke-RestMethod 'https://www.yammer.com/api/v1/messages.json' -Headers $headers -Method GET
        $messages = $response.messages 
        
        #If page = 1 we return the first request
        if ($page -eq 1){ 
            Return $messages           
        }
        else { 
            #If page is 2 or larger we need to loop through 
            for ($i = 2; $i -le $page){
                #Checking oldest message ID
                $ArrLength = ($messages.Length) -1
                $oldestMessage = $messages[$ArrLength].id
                
                #Creating target with the oldest message ID; Updating messages to include the latest request.
                $target = 'https://www.yammer.com/api/v1/messages.json?older_than=' + $oldestMessage
                $response = Invoke-RestMethod $target -Headers $headers -Method GET
                $messages = $response.messages
                                         
                #Checking if there are additional pages;
                #If no older_available stop the loop                                             
                if ($response.meta.older_available){
                    $i += 1                                
                }
                else {
                    $i = $page+1
                    $messages = ''
                    Write-Host 'No messages under page' $page                    
                }
            }
        }
     Return $messages   
    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
    }

}

function Get-YammerMessage
(
    [String]$messageID,
    [string]$olderThan,
    [string]$newerThan,
    [boolean]$threaded,
    [string]$limit,
    [string]$threadID,
    [string]$groupID,
    [Int]$page
){

<#
        .DESCRIPTION
        The main function to retrieve Yammer messages.

       
        .EXAMPLE
         Get-YammerMessage -messageID 744466867
         Get-YammerMessage -olderThan 744466867 -newerThan 743070986  -threaded $true -limit 2
         Get-YammerMessage -groupID 8667682
         Get-YammerMessage -page 2
         

        .NOTES

        #>

        if ($messageID){$messages = Get-YammerMessageByID -messageID $messageID}
        
        elseif (($olderThan) -or ($newerThan) -or ($threaded) -or ($limit)){$messages =  Get-YammerMessages -olderThan $olderThan -newerThan $newerThan -threaded $threaded -limit $limit }
        
        elseif($threadID){$messages = Get-YammerMessagesByThreadID -threadID $threadID }
        
        elseif($groupID){$messages =  Get-YammerMessagesByGroupID -groupID $groupID }
        
        elseif($page){$messages =  Get-YammerMessagesByPage -page $page }
        
        # The following will return top 20 messages BUT this should be executed with the page par (page =1)
        #if(($messageID.Length -eq 0) -and ($olderThan.Length -eq 0) -and ($newerThan.Length -eq 0) -and ($threaded.Length -eq 0) -and ($limit.Length -eq 0) -and ($threadID.Length -eq 0) -and ($groupID.Length -eq 0) -and ($page.Length -eq 0)) {$messages = Get-YammerMessages}

        Return $messages
}



function New-YammerMessage
(
    [Parameter(Mandatory=$true)][string]$messageBody,
    [string]$groupID,
    [string]$repliedToID,
    [string]$directToID,
    #[boolean]$announcement,
    [string]$announcementTitle,
    [string[]]$topic
    
){
<#
        .DESCRIPTION
        Create a new message in Yammer. Message body is mandatory. 
        Topic input format should be "AA", "BBB", "CC"

       
        .EXAMPLE
         New-YammerMessage -messageBody "BB" -groupID 8667682 -repliedToID 744475784
         New-YammerMessage -messageBody "aa" -directToID 1583661588 
         

        .NOTES

        #>

    
    try {
      
            $requestBody = @{}
            $requestBody.body = $messageBody
                if ($groupID){ 
                    $requestBody.group_id = $groupID 
                        if ($directToID) {
                                        Write-Host 'DirectedToID and GroupID cannot be execute in the same request; message will be posted in the group'
                                        $directToID = ""
                        }
                        if ($repliedToID) {
                                        $message = Get-YammerMessage -messageID $repliedToID
                                        if ($message.group_id -ne $groupID) {
                                            Write-Host 'The message' $repliedToID 'is not part of group' $groupID '. message will be posted in the group'.
                                            $repliedToID = ""
                                        }
                        }
                    }
                if ($directToID){ $requestBody.direct_to_id = $directToID } 
                if ($topic.Length -gt 0) { 
                                            $topicNum = 1;
                                            foreach ($tag in $topic){
                                                $topicN = 'topic'+$topicNum
                                                $requestBody.$topicN = $tag
                                                $topicNum +=1
                                            }
                                        }
                if ($repliedToID) { $requestBody.replied_to_id = $repliedToID }
                if ($announcementTitle) {
                                       $requestBody.message_type = 'announcement'
                                       $requestBody.title = $announcementTitle
                                   }
                          
            $requestBody = ConvertTo-Json $requestBody
            $requestBody
            $target = 'https://www.yammer.com/api/v1/messages.json'
            $response = Invoke-RestMethod $target -Headers $headers -Method POST -Body $requestBody -ContentType application/json
    } 
    
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
    }


    Return $response.messages

    }
    
    

function Remove-YammerMessage([string]$messageID,[boolean]$purge){
<#
        .DESCRIPTION
        Delete Yammer messages by message ID

       
        .EXAMPLE
       Remove-YammerMessage -messageID 123345
         

        .NOTES

        #>



      try {

            $target = 'https://www.yammer.com/api/v1/messages/' +$messageID
            
            if ($purge){
                $target += '?purge=true'
            }

            $response = Invoke-RestMethod $target -Headers $headers -Method DELETE -ContentType application/json
            Write-Host $messageID 'deleted successfully'

    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value__
    }


}

Export-ModuleMember -Function Get-YammerMessage
Export-ModuleMember -Function New-YammerMessage
Export-ModuleMember -Function Remove-YammerMessage
