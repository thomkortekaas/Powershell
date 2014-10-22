#
# title dotDesktop Exchange Tools v1.0
#
# This script can create, modify and remove tenants.
# It can create modify and remove tenants as well.
#
# 
# For the script to run properly the following features have to be installed
# 1. Active directory Powershell CMDLets
# 2. Group Policy Management
# 3. Powershell 4.0 (WMF Framework 4.0)
#
# 
# Created   sept 2014
# By        Thom Kortekaas
#



clear

#region load cmlets

# Copyright (c) Microsoft Corporation. All rights reserved.  

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://dxshdcas01/PowerShell/ -Authentication Kerberos 
Import-PSSession $Session -AllowClobber

Import-Module ActiveDirectory
Import-Module grouppolicy

#endregion

#region functions

Function Add-Tenant{
    #$TenantName = 
	#$TenantCode = 
	#$TenantUPN  = 

    $TenantName = Read-Host  "Tenant Name"
    $TenantCode = Read-Host  "Tenant Code"
    $TenantUPN  = Read-host  "Tenant UPN"

    Write-host "en de tenant is $TenantName"
    Write-host "en de tenant code is $TenantCode"
    Write-host "en de tenant upn is $TenantUPN"

    
	#Create OU
	New-ADOrganizationalUnit -Name $TenantName -Path "OU=Tenants,DC=cloud,DC=dotdesktop,DC=nl" -Description $TenantCode # -Verbose
    start-sleep -s 30

	#Create UPN
	Set-ADForest -Identity cloud.dotdesktop.nl -UPNSuffixes @{ add = $TenantUPN } # -Verbose

	#Create accepted domain
	New-AcceptedDomain -Name $TenantName -DomainName $TenantUPN
	
    #Create GAL
	New-GlobalAddressList -Name ($TenantName + " – GAL") -ConditionalCustomAttribute1 $TenantCode -IncludedRecipients MailboxUsers -RecipientContainer ("cloud.dotdesktop.nl/Tenants/" + $TenantName) # -Verbose
	
    #Create All Rooms Address list
	New-AddressList -Name  "$TenantName – All Rooms" -RecipientFilter "(CustomAttribute1 -eq `'$TenantCode`') -and (RecipientDisplayType -eq 'ConferenceRoomMailbox')" -RecipientContainer "cloud.dotdesktop.nl/Tenants/$TenantName" # -Verbose
	
	#Create All Users list
	New-AddressList -Name ($TenantName + " – All Users") -RecipientFilter "(CustomAttribute1 -eq '$TenantCode') -and (ObjectClass -eq 'User')" -RecipientContainer ("cloud.dotdesktop.nl/Tenants/" + $TenantName)  # -Verbose
	
    #Create all contacts list
	New-AddressList -Name ($TenantName + " – All Contacts") -RecipientFilter "(CustomAttribute1 -eq '$TenantCode') -and (ObjectClass -eq 'Contact')" -RecipientContainer ("cloud.dotdesktop.nl/Tenants/" + $TenantName) # -Verbose
	
    #Create all groups list
	New-AddressList -Name ($TenantName + " – All Groups") -RecipientFilter "(CustomAttribute1 -eq '$TenantCode') -and (ObjectClass -eq 'Group')" -RecipientContainer ("cloud.dotdesktop.nl/Tenants/" + $TenantName) # -Verbose
	
    #Create OAB
	New-OfflineAddressBook -Name $TenantName -AddressLists ($TenantName + " – GAL")  # -Verbose

    start-sleep -s 10
	
    #Create Address Book Policy
	New-AddressBookPolicy -Name $TenantName -AddressLists ($TenantName + " – All Users"), ($TenantName + " – All Contacts"), ($TenantName + " – All Groups") -GlobalAddressList ($TenantName + " – GAL") -OfflineAddressBook $TenantName -RoomList ($TenantName + " – All Rooms") # -Verbose
	
    #Create Email address policy
    New-EmailAddressPolicy -Name $TenantUPN -includedRecipients AllRecipients -RecipientContainer "cloud.dotdesktop.nl/Tenants/$TenantName"  -EnabledPrimarySMTPAddressTemplate "SMTP:%1g.%s@$tenantupn"  
    Start-Sleep -s 10
	Get-EmailAddressPolicy $TenantUPN | Update-EmailAddressPolicy

    #Create AD group for tenant
    New-ADGroup -Name "SG_AllUsers_$TenantName" -DisplayName "SG_AllUsers_$TenantName"  -Path "OU=$TenantName,OU=Tenants,DC=cloud,DC=dotdesktop,DC=nl" -GroupScope Global -GroupCategory Security | Set-ADObject -ProtectedFromAccidentalDeletion
    
    #takes some thing to replicatie new AD group to other DC's, 30 seconds delay should do the trick before applying group to folders
    start-sleep -s 30

    #Create Data folders structure for customer 
    #Set variables for (sub)folders
    $data               = "\\cloud.dotdesktop.nl\data\$TenantName"
    $dataCompanydata    = "\\cloud.dotdesktop.nl\data\$TenantName\CompanyData"
    $dataUserdata       = "\\cloud.dotdesktop.nl\data\$TenantName\UserData" 
    $dataUserdataRF     = "\\cloud.dotdesktop.nl\data\$TenantName\UserData\RedirectedFolders"
    $dataUserDataTS     = "\\cloud.dotdesktop.nl\data\$TenantName\Userdata\TSProfiles" 

    #Create folder structure
    New-Item -Path       $data                   -ItemType Directory
    New-Item -Path       $dataCompanydata        -ItemType Directory
    New-Item -Path       $dataUserdata           -ItemType Directory
    New-Item -Path       $dataUserdataRF         -ItemType Directory
    New-Item -Path       $dataUserDataTS         -ItemType Directory

    #Disable inheritance on all folders
    .\icacls.exe $data             /inheritance:d
    .\icacls.exe $dataCompanydata  /inheritance:d
    .\icacls.exe $dataUserdata     /inheritance:d
    .\icacls.exe $dataUserdataRF   /inheritance:d
    .\icacls.exe $dataUserDataTS   /inheritance:d

    #Remove default users group from folders, /T means recursive
    .\icacls.exe $data /remove:g users /T

    #Add usergroup to root with read permissions on this folder only
    .\icacls.exe $data /grant:r cloud\SG_AllUsers_$TenantName`:R

    #Add usergroup to companydata folder with modify rights
    .\icacls.exe $dataCompanydata /grant:r cloud\SG_AllUsers_$TenantName`:`(OI`)`(CI`)M

    #Add Usergroup to userdata folder
    .\icacls.exe $dataUserdata /grant:r cloud\SG_AllUsers_$TenantName`:R

    #$TenantName = "koudasfalt"
    #$dataUserdataRF = "\\cloud.dotdesktop.nl\data\Koudasfalt\Userdata\RedirectedFolders"
    #Add usergroup to redirected folders, special permissions. AD (append data/add subdirectory) & RD (read data/list directory) - This folder only
    .\icacls.exe $dataUserdataRF /grant:r cloud\SG_AllUsers_$TenantName`:`(RD`,AD`)
        
    #Add usergroup to TSProfiles folders, special permissions. AD (append data/add subdirectory) & RD (read data/list directory) - This folder only
    .\icacls.exe $dataUserDataTS /grant:r cloud\SG_AllUsers_$TenantName`:`(RD`,AD`)

    
    #Create group policies (empty)
    
    New-GPO -Name "USER - $TenantName Drive Mappings "     | New-GPLink -Target "OU=$TenantName,OU=Tenants,DC=cloud,DC=dotdesktop,DC=nl"
    New-GPO -Name "USER - $TenantName Folder Redirection " | New-GPLink -Target "OU=$TenantName,OU=Tenants,DC=cloud,DC=dotdesktop,DC=nl"



	}

Function Remove-Tenant {

$tenants = Get-ADOrganizationalUnit -SearchBase "OU=Tenants,DC=cloud,DC=dotdesktop,DC=nl" -SearchScope Subtree -Filter * -properties Description | ? { $_.Name -notmatch "Tenants" } | select Name,Description
$Removetenants = $tenants | Out-GridView -Title "Select tenants you want to remove permanently" -OutputMode Multiple 

#First delete Addressbook policy before removing Addresslist, because the addresslists are in use by the addressbook policy

$Removetenants | ForEach-Object{

    #Set variable for tenant name
    $RemoveTenantName = $_.name
    $RemoveTenantCode = $_.Description

    #Remove Addressbook policy
    Get-AddressBookPolicy | ?{$_.name -eq $RemoveTenantName} | Remove-AddressBookPolicy -Confirm:$false
    #Remove addresslists
    Get-AddressList | ?{$_.name -eq $RemoveTenantName} | Remove-AddressList -Confirm:$false
    #Remove Global Address List
    Get-GlobalAddressList | ?{$_.name -eq $RemoveTenantName} | Remove-GlobalAddressList -Confirm:$false
    #Remove accepted domain
    Get-AcceptedDomain | ?{$_.Name -eq $RemoveTenantName } | Remove-AcceptedDomain -Confirm:$false
    #Remove UPN suffix
    Set-ADForest -Identity cloud.dotdesktop.nl -UPNSuffixes @{Remove=$RemoveTenantName}
    #Remove Organization unit
    Set-ADOrganizationalUnit -Identity "$RemoveTenantName,OU=Tenants,DC=cloud,DC=dotdesktop,DC=nl" -ProtectedFromAccidentalDeletion:$false
    Remove-ADOrganizationalUnit -Identity "$RemoveTenantName,OU=Tenants,DC=cloud,DC=dotdesktop,DC=nl" -Confirm:$false



}

#Get-AddressBookPolicy | Where-Object {$_.name -match }

}

Function View-Tenants {

$tenants = Get-ADOrganizationalUnit -SearchBase "OU=Tenants,DC=cloud,DC=dotdesktop,DC=nl" -SearchScope Subtree -Filter * -properties Description | ? { $_.Name -notmatch "Tenants" } | select Name,Description
$tenants | Out-GridView -Title Tenants  # -PassThru
}

Function Modify-Tenant {
Write-host "lalallalal"
$tenants = Get-ADOrganizationalUnit -SearchBase "OU=Tenants,DC=cloud,DC=dotdesktop,DC=nl" -SearchScope Subtree -Filter * -properties Description | ? { $_.Name -notmatch "Tenants" } | select Name,Description
$tenants | Out-GridView -Title Tenants  -OutputMode Single 

}


#endregion




Do {
Write-Host "
---------- dotXS Hosted Desktop Tools v1.0 ----------
1 = Add Tenant
2 = Remove Tenant
3 = View Tenants
4 = Modify Tenant
5 = Add mailbox
6 = Remove mailbox
7 = View mailboxes
Q = Quit
--------------------------"
$choice1 = read-host -prompt "Select number & press enter"
} until ($choice1 -lt 7)

Switch ($choice1) {
"1" {Add-Tenant}
"2" {Remove-Tenant}
"3" {View-tenants}
"4" {Modify-Tenant}
"5" {Write next sub-menu here}
"6" {Write next sub-menu here}
"7" {Write next sub-menu here}
"q" {Write-host "Quit"}
}