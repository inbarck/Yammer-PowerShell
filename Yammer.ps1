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
<#
.SYNOPSIS
  Yammer Module for Windows PowerShell v 1.0
.DESCRIPTION
  The PowerShell Script Yammer.ps1 holds the global parameter $headers
  The script imports additional PowerShell Modules
.PARAMETER <Parameter_Name>

.INPUTS

.OUTPUTS

.NOTES
  Version:        1.0
  Author:         Inbar Cizer Kobrinsky
  Creation Date:  August 2016
  Purpose/Change: Wraps Yammer API with PowerShell
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

#Dot Source required Function Libraries
. "C:\Scripts\Functions\Logging_Functions.ps1"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.0"

#Log File Info
$sLogPath = "C:\Windows\Temp"
$sLogName = "<script_name>.log"
$sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName


#-----------------------------------------------------------[Execution]------------------------------------------------------------


$global:headers = $null 


## Import additional Modules
Import-Module .\YammerAuth.psm1 -Verbose
Import-Module .\YammerUsers.psm1 -Verbose
Import-Module .\YammerMessages.psm1

## In order to authenticate use Connect-YammerService

