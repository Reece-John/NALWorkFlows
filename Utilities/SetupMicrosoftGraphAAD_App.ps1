#SetupMicrosoftGraphAAD_App.ps1

*************************************************************
*************************************************************
*************************************************************
*************************************************************
*************************************************************
*************************************************************
*************************************************************
*********************   Stop this is not tested  ************
*************************************************************
*************************************************************
*************************************************************
*************************************************************

Exit;



connect-Azuread

$tenantID=(Get-AzureADTenantDetail).ObjectId
#client Secret is the application password . You can cenerate this as long as you dont have
# special characters like + or / that can prevent correct authentication 
$client_secret = "2Uban4QXXXXXqg6mcXdpXXX00zsTPPixJf6kXgcml3E="
$applicationName="PSTestGraphApp"
#you canuse any valid URL here
$homePage = "https://gmarculescu.com"
$appIdURI = "https://gmarculescu.com/?p=584"
$logoutURI = "http://portal.office.com"

#We create the application secret valid for one year starting today
$today=[System.DateTime]::Now
$keyId = (New-Guid).ToString();
$applicationSecret = New-Object Microsoft.Open.AzureAD.Model.PasswordCredential($null, $today.addyears(1), $keyId, $today, $client_secret)
# We create the AAD aplication 
$AADApplication = New-AzureADApplication -DisplayName $applicationName `
        -HomePage $homePage `
        -ReplyUrls $homePage `
        -IdentifierUris $appIdURI `
        -LogoutUrl $logoutURI `
        -PasswordCredentials $applicationSecret

# We create a service principal for our application application 
$servicePrincipal = New-AzureADServicePrincipal -AppId $AADApplication.AppId