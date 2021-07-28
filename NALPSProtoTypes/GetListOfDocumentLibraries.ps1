# GetListOfDocumentLibraries.ps1



#Set Variables
Clear-Host;

$tenantAbbreviation = "ALGA";

# get tenant specific variable values
$tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;


# get administrator credentials
[system.Management.Automation.PSCredential]$psAdminCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

#$SiteURL = "https://crescent.sharepoint.com/sites/marketing"
$SiteURL = "https://algeorgetownarea-my.sharepoint.com/personal/data_algeorgetownarea_org";
  
#Connect to PNP Online
Connect-PnPOnline -Url $SiteURL -Credentials $psAdminCredentials;
 
#Get all document libraries - Exclude Hidden Libraries
$DocumentLibraries = Get-PnPList | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $false} #Or $_.BaseType -eq "DocumentLibrary";
 
#Get Document Libraries Name, Default URL and Number of Items
$DocumentLibraries | Select-Object Title, DefaultViewURL, ItemCount;