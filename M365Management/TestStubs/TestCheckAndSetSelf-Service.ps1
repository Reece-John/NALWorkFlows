
#TestCheckAndSetSelf-Service.ps1

Clear-Host;

[string]$tenantAbbreviation = "ALGA";

# get tenant specific variable values
$tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;

# get administrator credentials
[System.Management.Automation.PSCredential]$tenantCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

#connect to SharePoint On-line
Connect-MsolService -Credential $tenantCredentials;



Get-MsolCompanyInformation | fl AllowAdHocSubscriptions


#Disable Self-Service
Set-MsolCompanySettings -AllowAdHocSubscriptions $false;


#enable Self-Service
#Set-MsolCompanySettings -AllowAdHocSubscriptions $true;
