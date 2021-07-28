#
# TestConnectToExchange.ps1
#


Clear-Host;

$startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");

Write-Host($startLoc);

Set-Location $startLoc;


$tenantAbbreviation = "ALGA";
$tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 $tenantAbbreviation;
$connectionUri = "https://outlook.office365.com/powershell-liveid/";

[system.Management.Automation.PSCredential]$psAdminCredentials = .\Common\ReturnTenantCredentials.ps1 $tenantAbbreviation $tenantObj;

#connect to Exchange Online
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $psAdminCredentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking
#$Perms | foreach {Get-ManagementRoleAssignment -Role $_.Name -Delegating $false | Format-Table -Auto Role,RoleAssigneeType,RoleAssigneeName}
#Get-mailbox it3@ALAustin.org | FL CustomAttribute1

$mb = Get-mailbox "mjohn@ALGeorgetownArea.org";
Write-Host($mb);
<#
$mb | Select-Object -Property *

Write-Host("MB CustomAttribute1: " + $mb.CustomAttribute1);
Write-Host("MB EmailAddresses: " + $mb.EmailAddresses);
Write-Host("MB UserPrincipalName: " + $mb.UserPrincipalName);

#$user = Get-MsolUser -UserPrincipalName it3@AlAustin.org


#$xx = $user."VolgisticsId";
#$user.ExtensionData
#Write-Host("VolgisticsId: " + $xx);

#Write-Host("UserPrincipalName: " + $user.UserPrincipalName + " - UserObjectId: " + $user.ObjectId);






#Get-MsolUser -UserPrincipalName it3@AlAustin.org | Format-List

#>
Remove-PSSession $Session




