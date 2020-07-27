#Retention Policies for Teams Chat: https://docs.microsoft.com/en-us/microsoftteams/retention-policies#:~:text=%20To%20create%20a%20retention,data,%20delete%20it,...%20More

#Install EXO management module
Install-Module -Name ExchangeOnlineManagement

#Log in
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

Get-RetentionCompliancePolicy -Identity 'Chat - Private - Retention Policy' -DistributionDetail | Format-List
Set-RetentionCompliancePolicy -Identity 'Chat - Private - Retention Policy' -AddTeamsChatLocation "254790cd-0d17-4cc0-a5d8-40d9dc126958","Hybrid User 1","hybriduser2@mebajric.msftonlinerepro.com","98304120-cbb3-4aec-849d-2920da33f229"
$nextuser = '8bbb5533-2218-410e-ba32-3052700ad7c2'
Set-RetentionCompliancePolicy -Identity 'Chat - Private - Retention Policy' -AddTeamsChatLocation $nextuser
$rCP = Get-RetentionCompliancePolicy -Identity 'Chat - Private - Retention Policy' -DistributionDetail | Format-List  
$rCP.get

#This will set the policies for specified users:
Get-RetentionCompliancePolicy -Identity 'Chat - Private - Retention Policy' -DistributionDetail | select TeamsChatLocation, @{TeamsChatLocation='TeamsChatLocation';"Expression"={$_.TeamsChatLocation -join ","}}
Get-RetentionCompliancePolicy -Identity 'Chat-Retention Policy' -DistributionDetail | select -ExpandProperty TeamsChatLocation | foreach {$_.DisplayName} | Sort-Object $_DisplayName

#Remove-RetentionCompliancePolicy:https://docs.microsoft.com/en-us/powershell/module/exchange/remove-retentioncompliancepolicy?view=exchange-ps
Remove-RetentionCompliancePolicy -Identity "Chat - Private - Retention Policy"

#New-RetentionCompliancePolicy https://docs.microsoft.com/en-us/powershell/module/exchange/new-retentioncompliancepolicy?view=exchange-ps
New-RetentionCompliancePolicy -Name "Chat-Retention Policy" -TeamsChatLocation "254790cd-0d17-4cc0-a5d8-40d9dc126958"

Get-RetentionCompliancePolicy -Identity 'Chat-Retention Policy' -DistributionDetail | Format-List
Get-RetentionCompliancePolicy -Identity 'Rar2' -DistributionDetail | Format-List

Get-RetentionComplianceRule -Identity 'Chat - Private - Retention Policy' | Format-List
Get-RetentionComplianceRule -Identity 'Rar' | Format-List

#New-RetentionComplianceRule: https://docs.microsoft.com/en-us/powershell/module/exchange/new-retentioncompliancerule?view=exchange-ps
New-RetentionComplianceRule -Name SeptOneYear -Policy "SLT" -RetentionDuration Unlimited


#When you create a policy in the GUI and select time/retention period, you can retrieve the settings in PS using the same identity from Get-RetentionPolicyCommand
Get-RetentionCompliancePolicy -Identity 'Rar' -DistributionDetail | Format-List
Get-RetentionComplianceRule -Identity 'Rar2' | Format-List

#First make compliance retention policy
New-RetentionCompliancePolicy -Name "chatRetention" -TeamsChatLocation "254790cd-0d17-4cc0-a5d8-40d9dc126958"

#Then, make the retention rule
#New-RetentionComplianceRule: https://docs.microsoft.com/en-us/powershell/module/exchange/new-retentioncompliancerule?view=exchange-ps
#The retention rule must be added to an existing retention policy using the Policy parameter. Only one rule can be added to each retention policy.
New-RetentionComplianceRule -Policy "chatRetention" -RetentionComplianceAction KeepAndDelete -ExpirationDateOption CreationAgeInDays 

#Add users from text file to the Existing chatRetention Policy:
$nextuser = '82b13ad2-312b-4b5d-bb41-2f3e97b3b30b'
Set-RetentionCompliancePolicy -Identity 'chatRetention' -AddTeamsChatLocation $nextuser

#Get 
Get-RetentionCompliancePolicy -Identity 'chatRetention' -DistributionDetail | Format-List
Get-RetentionComplianceRule -Policy "chatRetention" | Format-List

Get-RetentionCompliancePolicy -Identity 'chatRetention' -DistributionDetail | select -ExpandProperty TeamsChatLocation | foreach {$_.DisplayName} | Sort-Object $_.DisplayName 

$dispName = Get-RetentionCompliancePolicy -Identity 'chatRetention' -DistributionDetail | select -ExpandProperty TeamsChatLocation | foreach {$_.DisplayName}  
    
#checked perms per https://docs.microsoft.com/en-us/microsoft-365/security/office-365-security/permissions-in-the-security-and-compliance-center?view=o365-worldwide

#Output policy settings to a text file:
Get-RetentionCompliancePolicy -Identity 'chatRetention' -DistributionDetail | Format-List >"C:\Users\mebajric\Documents\Scripts and Templates\Teams\chatRetention-policy.txt"
Get-RetentionComplianceRule -Policy "chatRetention" | Format-List >"C:\Users\mebajric\Documents\Scripts and Templates\Teams\chatRetention-rule.txt"