﻿#
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
#$siteUrls = Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/" | select Url 

$site = Get-SPOSite https://melomel-my.sharepoint.com/personal/geowash_melomel_onmicrosoft_com | select Url 
$SPOCreds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($creds.UserName, $creds.Password)

$logFile = [Environment]::GetFolderPath("Desktop") + "\OneDriveSites-SingleUser.log"

$date = Get-Date -Format U
"========================================================================" | out-file $logFile -append
"Started Get-Remove Perms at " + $date + " (UTC time)" | out-file $logFile -append
"========================================================================" | out-file $logFile -append

$adminsPersonalSite = $creds.UserName.Replace('@','_').Replace('.','_')

#foreach($obj in $siteUrls){
    $url = $site.url
    Write-Host $url
    $url | out-file $logFile -append
   
    if(-not $url.Contains($adminsPersonalSite)){
        #Set the spouser as admin else it will get access denied on getting role assignments
        Set-SPOUser -site $url -LoginName $creds.UserName -IsSiteCollectionAdmin $True
    }
   
    #List Perms    
    Get-SPOUser -Site $url
    $listusers = Get-SPOUser -Site $url
    Write-Host "Original OneDrive perms: " + $listusers 
    "Original OneDrive perms: " | out-file $logFile -append
    $listusers | out-file $logFile -append
    
    #Remove Group - comment out these lines to add the group - update the guid with the guid of your everyone except external users guid
    Remove-SPOUser -Site $url -LoginName "c:0-.f|rolemanager|spo-grid-all-users/a6b27fb1-cd85-4932-bc29-05ab5890a64a" -Group $group.LoginName
    #Remove-SPOUser -Site $url -LoginName "c:0t.c|tenant|1b92dd5d-6cff-429a-9a77-79de8b445beb" -Group $group.LoginName
    Write-Host "Removed group" 
    "Removed group" | out-file $logFile -append
    " " | out-file $logFile -append

    #Add Group - this is wonky/not right but looks to add the group??
    #Set-SPOUser -Site $url -LoginName "c:0t.c|tenant|1b92dd5d-6cff-429a-9a77-79de8b445beb" 
    #"Added group" | out-file $logFile -append
    #" " | out-file $logFile -append
   

    #List Perms after removing group
    Get-SPOUser -Site $url
    $usersAfter = Get-SPOUser -Site $url
    Write-Host "OneDrive permissions after removing everyone except external users: " + $usersAfter 
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
#}



#followups: 
#Run on mysite? Example 4 prevents non-owners of a site from inviting new users to the site.
#https://docs.microsoft.com/en-us/powershell/module/sharepoint-online/set-sposite?view=sharepoint-ps
#Set-SPOSite -Identity https://contoso.sharepoint.com -DisableSharingForNonOwners
