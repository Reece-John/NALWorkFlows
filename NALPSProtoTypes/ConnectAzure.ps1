# ConnectAzure.ps1



#https://devblogs.microsoft.com/premier-developer/azure-active-directory-automating-guest-user-management/



#https://itfordummies.net/2020/11/29/microsoft-graph-api-powershell-azuread-app/

# completed up to add certificate

#https://docs.microsoft.com/en-us/powershell/azure/authenticate-azureps?view=azps-5.7.0


<#
Create a self-signed root certificate
Use the New-SelfSignedCertificate cmdlet to create a self-signed root certificate.
For additional parameter information, see New-SelfSignedCertificate.

From a computer running Windows 10 or Windows Server 2016, open a Windows PowerShell console with elevated privileges. 
These examples do not work in the Azure Cloud Shell "Try It". You must run these examples locally.

Use the following example to create the self-signed root certificate. The following example creates a self-signed root certificate
 named 'P2SRootCert' that is automatically installed in 'Certificates-Current User\Personal\Certificates'.
 You can view the certificate by opening certmgr.msc, or Manage User Certificates.

Sign in using the Connect-AzAccount cmdlet. Then, run the following example with any necessary modifications.
#>


Clear-Host;

$machineUsage = [Environment]::GetEnvironmentVariable("MachineUsage","Machine");
if($machineUsage -ne "Production")
{
    $startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
    Set-Location $startLoc;
}

[string]$tenantAbbreviation = "NAL";

# get tenant specific variable values
$tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;

# get administrator credentials
[system.Management.Automation.PSCredential]$psAdminCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;


Connect-AzureAD -TenantDomain "ALGeorgetownarea.org" -Credential $psAdminCredentials
#Connect-AzureAD -Credential $psAdminCredentials

Disconnect-AzureAD;

<#
$cert = New-SelfSignedCertificate -Type Custom -KeySpec Signature `
-Subject "CN=P2SRootCert" -KeyExportPolicy Exportable `
-HashAlgorithm sha256 -KeyLength 2048 `
-CertStoreLocation "Cert:\CurrentUser\My" -KeyUsageProperty Sign -KeyUsage CertSign
#>




Write-Host("Disconnected");



<#

$sp = New-AzADServicePrincipal -DisplayName ServicePrincipalName

# Retrieve the plain text password for use with `Get-Credential` in the next command.
$sp.secret | ConvertFrom-SecureString -AsPlainText

$pscredential = Get-Credential -UserName $sp.ApplicationId
Connect-AzAccount -ServicePrincipal -Credential $pscredential -Tenant $tenantId

$pscredential = New-Object -TypeName System.Management.Automation.PSCredential($sp.ApplicationId, $sp.Secret)
Connect-AzAccount -ServicePrincipal -Credential $pscredential -Tenant $tenantId

#>