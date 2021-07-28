#
# TestGetTodaysLatestCHUserExportFile.ps1
#
Clear-Host;

$startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
Set-Location $startLoc;

[string]$tenantAbbreviation = "ALGA";
$tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation;


[string]$domainName = $tenantObj.DomainName;
[string]$domainExtension = $tenantObj.DomainExtension;


# create the log file name
$dateRightNow = Get-Date;
[string]$myMasterLogFilePathAndName = 'c:\logs\' + $tenantAbbreviation + '\TestRenameAndMoveChapterHubCSVFile_' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

$SharePointSiteURL = "https://" + $domainName +"-my.sharepoint.com/personal/data_" + $domainName + "_" + $domainExtension;
$SharePointFileRelativeURL = "/Documents/" + $tenantAbbreviation + "%20Chapter%20Hub%20Exports/";
$ListName = "/Documents/" + $tenantAbbreviation + "%20Chapter%20Hub%20Exports/";
#$ListName = "/Documents/Chapter%20Hub%20Exports/";
$baseExportFileName = $tenantObj.UserExportFileName;
$baseExportFileExtension = "csv";

#Get Credentials to connect
[system.Management.Automation.PSCredential]$myCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

if($true)
{
    $fileNameFound = .\Utilities\GetTodaysLatestTenantCHUserExportFile.ps1 -tenantCredentials $myCredentials `
                                                                           -sharePointSiteURL $SharePointSiteURL `
                                                                           -sharePointFileRelativeURL $SharePointFileRelativeURL `
                                                                           -sharePointListName $ListName `
                                                                           -baseExportFileName $baseExportFileName `
                                                                           -exportFileExtension $baseExportFileExtension `
                                                                           -masterLogFile $myMasterLogFilePathAndName;
}
Write-Host($fileNameFound);
