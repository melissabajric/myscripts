<#
.SYNOPSIS
  Add users to Teams chat retention policy or add Teams to Teams channel messages retention policy via text file
    
.DESCRIPTION 
  Add users to Teams chat retention policy or add Teams to Teams channel messages retention policy. 
  This script MUST run as a user with Global Admin permission in the M365 Admin Portal
  This script is provided as-is with no warranties expressed or implied. 
  This script is a quick and dirty example of how one could update Teams chat and channel retention policies. 
  USE AT YOUR OWN RISK!
  
  
.NOTES
  Version:        2.0
  Author:         Melissa Bajric
  Blog:			  https://melissabajric.us
  Creation Date:  16 June 2020
  Purpose/Change: Initial script development
  Modified Date:  30 June 2020 by mebajric -Updated added functions and updates for Teams channel messages retention policy

.EXAMPLE
  Update Teams retention policies for channel and/or chat
  First update function variables and log files with proper path and teams or users
  Set the variables: $acctsFile, $logFile, $copmPolicyName, $upn

  .\UpdateTeamsRetention.ps1 
#>

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#-----------------------------------------------------------[Functions]------------------------------------------------------------

function addUsersChatRetentionPolicy {

#setup for logging: Log name reflect the policy
$acctsFile = "C:\Users\mebajric\Documents\Scripts and Templates\Teams\TeamsRetention-AddUsers.txt"
$logFile = "C:\Users\mebajric\Documents\Scripts and Templates\Teams\TeamsChatRentionUsers-AddUsersOrTeamsLog.txt" # The output log file

$newStreamReader = New-Object System.IO.StreamReader($acctsFile)

$counter = 0
$totalAccts = get-content ($acctsFile) | Measure-Object -Line
$totalAcctsCount = $totalAccts.Lines 

#Compliance Policy Name:
$compPolicyName = "chatRetention"
$upn = "mebajric@melomel.onmicrosoft.com"

#Update chat policy with individual users:
$date = Get-Date -Format U
"========================================================================" | out-file $logFile -append
"Started AddUsers at " + $date + " (UTC time)" | out-file $logFile -append
"========================================================================" | out-file $logFile -append

while (($readEachLine = $newStreamReader.ReadLine()) -ne $null)
{
    $nextuser = $readEachLine
         
    $counter++ 
   
   "Adding user: " + ($counter) + " out of " + $totalAcctsCount + " --> " + $nextuser | out-file $logFile -Append
   
    try
        {
        
        $policyVars = Get-RetentionCompliancePolicy -Identity $compPolicyName -DistributionDetail | select -ExpandProperty TeamsChatLocation 
            if ($_.DisplayName -like "All") {
            return
            }
        Set-RetentionCompliancePolicy -Identity $compPolicyName -AddTeamsChatLocation $nextuser -ErrorAction Stop
        write-host "Added user: " ($counter) " out of " $totalAcctsCount " --> " $nextuser
        }
      
    catch [system.Exception]
       {
        
       "ERROR!!! for user: " + $nextuser + " -- " + $_.Exception.Message | out-file $logFile -append
        write-host "ERROR!!! for user: " + $nextuser + " -- " + $_.Exception.Message $nextTeam
       }
    
   }
$newStreamReader.Dispose()

}


function addTeamsChannelRetentionPolicy { 

#setup for logging: Log name reflect the policy
$acctsFile = "C:\Users\mebajric\Documents\Scripts and Templates\Teams\TeamsRetention-AddTeams.txt"
$logFile = "C:\Users\mebajric\Documents\Scripts and Templates\Teams\TeamsChatRentionUsers-AddUsersOrTeamsLog.txt" # The output log file

$newStreamReader = New-Object System.IO.StreamReader($acctsFile)

$counter = 0
$totalAccts = get-content ($acctsFile) | Measure-Object -Line
$totalAcctsCount = $totalAccts.Lines 

#Compliance Policy Name:
$compPolicyName = "chatRetention"
$upn = "mebajric@melomel.onmicrosoft.com"
    
#Update channel policy with specific teams:
$date = Get-Date -Format U
"========================================================================" | out-file $logFile -append
"Started AddTeam at " + $date + " (UTC time)" | out-file $logFile -append
"========================================================================" | out-file $logFile -append

while (($readEachLine = $newStreamReader.ReadLine()) -ne $null)
{
    $nextTeam = $readEachLine
         
    $counter++ 
   
   "Adding team: " + ($counter) + " out of " + $totalAcctsCount + " --> " + $nextTeam | out-file $logFile -Append
  
    try
      {

      $policyVars = Get-RetentionCompliancePolicy -Identity $compPolicyName -DistributionDetail | select -ExpandProperty TeamsChannelLocation 
          if ($_.DisplayName -like "All") {
          return
          }
       Set-RetentionCompliancePolicy -Identity $compPolicyName -AddTeamsChannelLocation $nextTeam -ErrorAction Stop
       write-host "Added Team: " ($counter) " out of " $totalAcctsCount " --> " $nextTeam
       
        }
    catch [system.Exception]
     {
       "ERROR!!! for Team: " + $nextTeam + " -- " + $_.Exception.Message | out-file $logFile -append
        write-host "ERROR!!! for Team: " + $nextTeam + " -- " + $_.Exception.Message $nextTeam
        }
    
   }
$newStreamReader.Dispose()


}

#-----------------------------------------------------------[Begin Script]------------------------------------------------------------

#Logging in
#Log in without MFA
#Install-Module -Name ExchangeOnlineManagement
#$UserCredential = Get-Credential
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
#Import-PSSession $Session -DisableNameChecking

#Log in MFA
#First install EXO https://docs.microsoft.com/en-us/powershell/exchange/mfa-connect-to-exchange-online-powershell?view=exchange-ps
Connect-IPPSSession -UserPrincipalName $upn 

addUsersChatRetentionPolicy
addTeamsChannelRetentionPolicy

Get-PSSession | Remove-PSSession

#-----------------------------------------------------------[End Script]------------------------------------------------------------
