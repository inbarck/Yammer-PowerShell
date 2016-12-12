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

 -------- Authentication and headers ---------------------------------------


#####################################################################################################################################################>


########## Primary function for authentication; This function builds the headers

function Connect-YammerService(){

        Write-Host ""
        Write-Host 'Welcome to Yammer PowerShell Module'
        Write-Host '============================================================='
        Write-Host ''
        Write-Host ''
        Write-Host 'In order to run the PowerShell Module you will need to get an app token'
        Write-Host 'If you do not have an app - Start by creating an app in Yammer: https://developer.yammer.com/docs/getting-started'
        Write-host ''
        Write-host ''

        Write-Host 'As the app creator you can generate personal token for testing'
        $token = Read-Host 'Please provide your token'
        
     
        $tempHeaders = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
        $tempHeaders.Add('Authorization', 'Bearer ' + $token)
        
     
        $global:headers = $tempHeaders
        
}




function Get-YammerObject() { 
        <#
        .DESCRIPTION
        Checks to see whether a computer is pingable or not.

        .PARAMETER computername
        Specifies the computername.

        .EXAMPLE
        IsAlive -computername testwks01

        .NOTES
        This is just an example function.
        #>
  
    try {
        $target = 'https://www.yammer.com/api/v1/users/current.json'
       # $response = Invoke-RestMethod $target  -Headers $headers -Method GET
        $response = Invoke-WebRequest -Uri $target -Headers $headers -Method GET 
        $object =  ConvertFrom-Json $response.Content
    }
    Catch [system.exception] {
        Write-Host 'Error: Failed to execute request' $target
        Write-Host 'HTTP Status Code:' $_.Exception.Response.StatusCode.Value
        $_.Exception.Response.StatusCode.StatusDescription
    }

    Return $object
}


Export-ModuleMember -Function Connect-YammerService
