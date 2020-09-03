#
# Melissa Bajric
# 09/03/2020
# 
# Description: 
# $tenantAdminUrl - should be in the format https://tenantName-admin.sharepoint.com
# 
# This script will prompt for SPO global admin credentials the first run, and then iterate thru each ODB site and output
#
# **Must TEMPORARILY TURN OFF MFA FOR THIS TO WORK!!!!
#
# Code Example Disclaimer:
# Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED 'AS IS'
# -This is intended as a sample of how code might be written for a similar purpose and you will need to make changes to fit to your requirements. 
# -This code has not been tested.  This code is also not to be considered best practices or prescriptive guidance.  
# -No debugging or error handling has been implemented.
# -It is highly recommended that you FULLY understand what this code is doing  and use this code at your own risk.
#
#!!! Script has not been tested for performance or with high volume.  Use at your OWN RISK!!!
######################################################################################################

#Update with your Tenant Admin URL here:
$tenantAdminUrl = "https://melomel-admin.sharepoint.com" # Change to your Admin URL

if($creds -eq $null){
    $creds = get-credential -Message "Enter credentials:"
}
connect-sposervice -url $tenantAdminUrl -Credential $creds
 

$SPOCreds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)

#Update files - acctsFile is a list of OneDrives for which to execute the script - logFile is an output file 
$acctsFile = "C:\Users\mebajric\Documents\Scripts and Templates\SPO\MyOnedrives.txt"
$logFile = [Environment]::GetFolderPath("Desktop") + "\OneDriveSites-FromBatch.log"

#For reading in the accounts
$newStreamReader = New-Object System.IO.StreamReader($acctsFile)

$counter = 0
$totalAccts = get-content ($acctsFile) | Measure-Object -Line
$totalAcctsCount = $totalAccts.Lines 

$date = Get-Date -Format U
"========================================================================" | out-file $logFile -append
"Started Get-Remove Perms at " + $date + " (UTC time)" | out-file $logFile -append
"========================================================================" | out-file $logFile -append

$adminsPersonalSite = $creds.UserName.Replace('@','_').Replace('.','_')

while (($readEachLine = $newStreamReader.ReadLine()) -ne $null)
{
    $counter++
    try
    {

    $site = Get-SPOSite $readEachLine | select Url 
    $url = $site.url
    Write-Host $url
    $url | out-file $logFile -append
   
    if(-not $url.Contains($adminsPersonalSite)){
        #Set the spouser as admin else it will get access denied on getting role assignments
        Set-SPOUser -site $url -LoginName $creds.UserName -IsSiteCollectionAdmin $True
    }
   
    #List Perms    
    Get-SPOUser -Site $url
    $users = Get-SPOUser -Site $url
    Write-Host "Original OneDrive perms: " + $users 
    "Original OneDrive perms: " | out-file $logFile -append
    $users | out-file $logFile -append
    
    #Remove Group - Update the guid with the guid of your everyone except external users guid
    Remove-SPOUser -Site $url -LoginName "c:0-.f|rolemanager|spo-grid-all-users/a6b27fb1-cd85-4932-bc29-05ab5890a64a" -Group $group.LoginName
    Write-Host "Removed group" 
    "Removed group" | out-file $logFile -append
    " " | out-file $logFile -append
   

    #List Perms after removing group
    Get-SPOUser -Site $url
    $usersAfter = Get-SPOUser -Site $url
    Write-Host "OneDrive permissions after removing everyone except external users: " + $users 
    "OneDrive permissions after removing everyone except external users" | out-file $logFile -append
    $usersAfter | out-file $logFile -append

    $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($url)
    $ctx.Credentials = $SPOCreds

    $w = $ctx.Web
    $ctx.Load($w)
    $ctx.ExecuteQuery()

    $users = $w.SiteUsers

    $ctx.Load($users)
    $ctx.ExecuteQuery()

    foreach ($user in $users) {
        if ($user.isSiteAdmin -eq $true) {
            Write-Output "$($user.LoginName) ($($user.Title))"
        }
    }
    
    if(-not $url.Contains($adminsPersonalSite)){
        #Now remove the Spo user permission
        Set-SPOUser -site $url -LoginName $creds.UserName -IsSiteCollectionAdmin $False
    }

Write-Host "Done with: " $url
"Done with: $($url) " | out-file $logFile -append
" " | out-file $logFile -append
    }
catch [system.Exception]
        {"ERROR!!! on OneDrive: " + $url + " -- " + $_.Exception.Message | out-file $logFile -append}
    
   }
$newStreamReader.Dispose() 
