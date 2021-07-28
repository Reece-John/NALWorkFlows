#
# TestRemoveArchiveFilesOlderThan.ps1
#
Clear-Host;

# Set relative Location
#$startLoc = [Environment]::GetEnvironmentVariable("PSStartup","Machine");
$startLoc = [Environment]::GetEnvironmentVariable("DevStartup","User");
Set-Location $startLoc;

$tenantAbbreviation= "NAL";

[PSCustomObject]$tenantObj = .\Common\ReturnTenantSpecificVariables.ps1 -tenantAbbreviation $tenantAbbreviation

[string]$domainName = $tenantObj.DomainName;
[string]$domainExtension = $tenantObj.DomainExtension;

# create the log file name
$dateRightNow = Get-Date;
[string]$myMasterLogFilePathAndName = 'c:\logs\' + $tenantAbbreviation + '\TestRemoveArchiveFilesOlderThan_' + $dateRightNow.ToString("yyyyMMddTHHmmss") + '.log';

#Get Credentials to connect
$myCredentials = .\Common\ReturnTenantCredentials.ps1 -tenantAbbreviation $tenantAbbreviation -tenantObj $tenantObj;

# Configuration Variables
[string]$archiveSiteURL = "https://" + $domainName +"-my.sharepoint.com/personal/data_" + $domainName + "_" + $domainExtension;
[string]$archiveRelativeURL = "/Documents/" + $tenantAbbreviation + "%20Chapter%20Hub%20Exports/";
#[string]$archiveRelativeURL = "/Documents/Chapter%20Hub%20Exports/";
[int]$daysOlderThan = 33;
[string[]]$extensionArray = @("csv", "log");

# Debug write statements
if($true)
{
    #region Debug write statements
    Write-Output($archiveSiteURL);
    Write-Output($archiveRelativeURL);
    Write-Output($daysOlderThan);
    Write-Output($extensionArray);
    #endregion Debug write statements
}

if($true)
{
    .\M365OneDrive\RemoveArchiveFilesOlderThan.ps1 -tenantCredentials $myCredentials `
                                                   -archiveSiteURL $archiveSiteURL `
                                                   -archiveRelativeURL $archiveRelativeURL `
                                                   -daysOlderThan $daysOlderThan `
                                                   -extensionArray $extensionArray `
                                                   -masterLogFilePathAndName $myMasterLogFilePathAndName;
}

Write-Host("Wrap Up TestRemoveArchiveFilesOlderThan.ps1");
