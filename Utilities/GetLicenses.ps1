#
# GetLicenses.ps1
#
<# Header
#*********************************
#           Author: Mike John
#     Date Created: 04/28/2021
#*********************************
# Last date edited: 04/28/2021
#   Last edited By: Mike John
# Last Edit Reason: Original
#*********************************
#>

#starts Here
Clear-Host;

$machineUsage = [Environment]::GetEnvironmentVariable("MachineUsage","Machine");
if($machineUsage -ne "Production")
{
    $startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
    Set-Location $startLoc;
}

[string]$tenantAbbreviation = "ALGA";

# get tenant specific variable values
$tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;

# get administrator credentials
[system.Management.Automation.PSCredential]$psAdminCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;


#connect to SharePoint On-line
Connect-MsolService -Credential $psAdminCredentials;


Get-MsolUser | where {$_.Licenses.AccountSkuId -like "*:SPB*"} | Format-List DisplayName,Licenses
#Get-MsolUser | where {$_.Licenses.AccountSkuId -like "*:*"} | Format-List DisplayName,Licenses
#Get-MsolUser | where {$_.Licenses.AccountSkuId -like "*:VISIOCLIENT*"} | Format-List DisplayName,Licenses

#Get-MsolUser | where {$_.Licenses.AccountSkuId -eq "algeorgetownarea:SPB"} | Format-List DisplayName,Licenses
#Write-Host("--------------------------------------------");
#Get-MsolUser | where {$_.Licenses.AccountSkuId -eq "algeorgetownarea:FLOW_FREE"} | Format-List DisplayName,Licenses
#Get-MsolUser -UserPrincipalName "bjohn@algeorgetownarea.org" | Format-List DisplayName,Licenses

<#
#Get-MsolUser | Format-List DisplayName,Licenses
#Get-MsolUser | where {$_.Licenses.AccountSkuId -eq "algeorgetownarea:SPB"} | Format-List DisplayName,Licenses
#Get-MsolUser | where {$_.AccountSkuId -eq "algeorgetownarea:SPB"} | Format-List DisplayName,Licenses

#(Get-MsolAccountSku | where {$_.AccountSkuId -like "*algeorgetownarea:SPB*"}).ServiceStatus
#>



