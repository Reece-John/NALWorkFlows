#
# TestConnectToSharePoint.ps1
#


# starts here
cls;

$startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
cd $startLoc;
$cred = .\Common\ReturnCredentials.ps1;


Get-SPOSite -Identity https://algeorgetownarea.sharepoint.com/sites/General

Get-SPOSite

Get-SPOSite -Identity https://algeorgetownarea.sharepoint.com/sites/ITInfrastructure
