#
# ManageSharedMailBoxs.ps1
#

<#
    ****** SERVICE ACCOUNT VS. APP REGISTRATION *****
    https://www.ravenswoodtechnology.com/authentication-options-for-automated-azure-powershell-scripts-part-1/
    https://www.ravenswoodtechnology.com/authentication-options-for-automated-azure-powershell-scripts-part-2/

    ****** Create Azure Certificate *****
    https://docs.microsoft.com/en-us/azure/active-directory/develop/howto-create-service-principal-portal

    https://docs.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps
    https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exo-v2-module

    ******* 2015 - Good explaination, but PowerShell part is outdated
    https://msexperttalk.com/understanding-shared-mailbox-limitations-office-365/#:~:text=Create%20a%20Shared%20Mailbox%20in%20Office%20365%20using,to%20view%20and%20send%20email%20from%20shared%20mailbox.
#>

function Connect-ExchOnline([string]$connectionUri, [system.Management.Automation.PSCredential]$tenantCredentials)
{
    #Create remote PowerShell session with Exchange On-line
    $ExchOnlineSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $tenantCredentials -Authentication Basic -AllowRedirection;

    #Import the remote PowerShell session
    Import-PSSession $ExchOnlineSession -AllowClobber | Out-Null;
}

<#
https://www.codetwo.com/kb/how-to-connect-to-exchange-server-via-powershell/#:~:text=To%20successfully%20connect%20to%20Exchange%20Online%20with%20PowerShell%2C,the%20-Credential%20parameter%20isn%E2%80%99t%20supported%20with%20MFA%20enabled.
#>
Clear-Host;

    $tenantAbbreviation = "ALGA";

    # get tenant specific variable values
    $tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;

    # get administrator credentials
    [system.Management.Automation.PSCredential]$psAdminCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

    [string]$tenantDomain = $tenantObj.DomainName;



$connectionUri = "https://" + $tenantDomain + "." + $tenantObj.DomainExtension + "/PowerShell";

$connectionUri = "https://outlook.office365.com/powershell-liveid/";
#$connectionUri = "https://algeorgetownarea.org/PowerShell";

#Connect-ExchangeOnline -Credentials $psAdminCredentials
#Connect-ExchangeOnline -UserPrincipalName navin@contoso.com
#Connect-ExchOnline -tenantCredentials $psAdminCredentials -connectionUri $connectionUri;

$mailBoxName = "Web Master";
$mailBoxAlias = "shared-TechTalk";
#$mailBoxTrustee = "riaz";

#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $psAdminCredentials -Authentication Basic -AllowRedirection
 $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $connectionUri -Credential $psAdminCredentials -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

$mb = Get-mailbox "mjohn@ALGeorgetownArea.org";
Write-Host($mb);


Remove-PSSession $Session

#New-Mailbox -Name $mailBoxName -Alias $mailBoxAlias -Shared;

#Add-RecipitentPermission $mailBoxName -Trustee $mailBoxTrustee -AccessRights SendAs;


<#

https://docs.microsoft.com/en-us/answers/questions/275036/exchange-online-shared-mailbox-bulk-create.html

 $Datas  = import-csv d:/temp/shared.csv
    
 foreach($Data in $Datas){
     if($null -eq (Get-Mailbox $data.name -erroraction 'silentlycontinue')){
         New-Mailbox -Name $data.Name -DisplayName $data.DisplayName -Shared
     }
     Add-MailboxPermission -Identity $data.Name -User $data.User -AccessRights $data.AccessRights
 }

#>

