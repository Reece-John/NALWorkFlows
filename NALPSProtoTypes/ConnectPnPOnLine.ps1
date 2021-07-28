#ConnectPnPOnLine.ps1

# This requires that an entry into the Windows Credential manager be created with global privileges 
# https://pnp.github.io/powershell/articles/authentication.html

Clear-Host;

#$tenantAbbreviation = "ALGA";
#$tenantAbbreviation = "ALSA";
$tenantAbbreviation = "ALLV";

# get tenant specific variable values
$tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;


# get administrator credentials
[system.Management.Automation.PSCredential]$psAdminCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

[string]$connectUrl = "https://" + $tenantObj.DomainName + ".sharepoint.com";


Connect-PnPOnline -Url $connectUrl -Credentials $psAdminCredentials

[string]$userIdentity = "";

if($tenantAbbreviation -eq "ALGA")
{
    $userIdentity ="mjohn@ALgeorgetownarea.org";
}

if($tenantAbbreviation -eq "ALSA")
{
    $userIdentity ="mjohn@ALSanAntonio.org";
}

if($tenantAbbreviation -eq "ALLV")
{
    $userIdentity ="mjohn@ALLV.org";
    #$userIdentity ="cwarman@ALLV.org";
}

Get-PnPAzureADUser -Identity $userIdentity;

Get-PnPAzureADUser -Filter "accountEnabled eq true"