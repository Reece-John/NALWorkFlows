#
# TestGetMailboxForwardingRules.ps1
#

Clear-Host;

$startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
Write-Host($startLoc);
cd $startLoc;

$uName = "MJohn";
$userDomain = "ALGeorgetownArea";
$domainExtension = ".org";
$userName = $uName + "@" + $userDomain + $domainExtension;

[system.Management.Automation.PSCredential]$tmpCreds = .\Common\GetUserCredentials.ps1 $uName $userDomain;
$clearTextPassword = $tmpCreds.GetNetworkCredential().password
$password = $clearTextPassword | ConvertTo-SecureString -asPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($username,$password);

#connect to Exchange Online
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking


#$mb = Get-mailbox mjohn@ALGeorgetownArea.org

Get-Mailbox -ResultSize Unlimited `
-Filter "ForwardingAddress -like '*' -or ForwardingSmtpAddress -like '*'" |
Select-Object Name,ForwardingAddress,ForwardingSmtpAddress;

Remove-PSSession $Session




