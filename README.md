# Yammer-PowerShell
Create a PowerShell script that wraps the Yammer REST API

** Important! This script is not supported by Microsoft or Yammer. ** 


Download all files and save them to the same folder. Open PowerShell and navigate to the above folder.

Run Yammer.ps1 – this script will load the Yammer PowerShell modules (psm1 files) - YammerAuth.psm1, YammerMessages.psm1 and YammerUsers.psm1. 

Run Connect-YammerService to authenticate and provide token for the app. 

If you don’t have an app – follow the documentation to create a Yammer app: https://developer.yammer.com/docs/getting-started. Once you have an app – go to the app page and click on “Generate a developer token for this application”. Use this token to authenticate in the PowerShell script.

Available commands:
Users: Get-YammerUser, New-YammerUser, Set-YammerUser  and Remove-YammerUser
Messages: Get-YammerMessage, New-YammerMessage, Remove-YammerMessage
