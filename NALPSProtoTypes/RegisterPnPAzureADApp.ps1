
[string]$ApplicationName = "TestApp";
[string]$FullTenantName = "yourtenant.onmicrosoft.com";
[string]$CertPathAndName = "c:\certificate.pfx";
[string]$CertPassword = "RequestedFromUser";
[string]$userTenantEmail = "mjohn@assistanceleague.org";
[string]$UserPassword = "RequestedFromUser";

$CertPassword = Read-Host -AsSecureString -Prompt "Enter Cert password";
$userTenantEmail = Read-Host -AsSecureString -Prompt "Enter User Tenant Email";
$UserPassword = Read-Host -AsSecureString -Prompt "Enter User password";

Register-PnPAzureADApp -ApplicationName $ApplicationName -Tenant $FullTenantName -CertificatePath $CertPathAndName -CertificatePassword $CertPassword -Username $userTenantEmail -Password $UserPassword;
